(() => {
  const ADVANCED_MODE = false;

  // =====================================================================
  // CONFIG — verify these against your actual Outlook Web interface.
  // Right-click the relevant button in DevTools > Inspect to confirm the
  // aria-label / text. These are the most common values as of Outlook
  // Web (outlook.office.com / outlook.live.com), but Microsoft changes
  // these periodically and they vary by locale/language.

  // Slower with "Extra safety delays"
  // Presses Escape and waits ~200ms before every right-click (to avoid stray menus)
  // Waits ~400ms between clicking menu-Delete and clicking confirm-Delete
  // Waits 500ms after that to check for a recurring-event dialog
  // Waits 1.5 seconds between each deleted event (vs. 0.8s in the Google Calendar original)
  // =====================================================================

  const SELECTORS = {
    // The "forward" arrow that advances the calendar view (month/week)
    nextPageButton: 'button[aria-label="Next"], button[aria-label*="next" i]',

    // The trash/delete icon that appears after opening an event.
    // Confirmed from DevTools: Outlook's delete button has NO aria-label
    // or title — it's a Fluent UI button identified only by its visible
    // text, "Delete". We match on text instead (see waitForButtonByText).
    deleteEventButtonText: 'Delete',

    // The dialog that pops up asking whether to delete just this occurrence
    // or the whole series (recurring events)
    recurringDialogSelector: 'div[role="dialog"], div.ms-Dialog-main',
    // Text that appears inside that dialog when it's the recurring-event prompt
    recurringDialogIndicatorText: ['occurrence', 'series', 'recurring'],
    // The button text inside that dialog to delete the single occurrence
    // (change to whatever the "delete series" button says if you want to
    // wipe the whole series instead — check with DevTools the same way)
    recurringDialogConfirmButtonText: 'Delete'
  };

  // Enhanced logging utility
  const logger = {
    info: (message, data = null) => {
      const timestamp = new Date().toISOString();
      console.log(`[INFO ${timestamp}] ${message}`, data || '');
    },
    warn: (message, data = null) => {
      const timestamp = new Date().toISOString();
      console.warn(`[WARN ${timestamp}] ${message}`, data || '');
    },
    error: (message, error = null) => {
      const timestamp = new Date().toISOString();
      console.error(`[ERROR ${timestamp}] ${message}`, error || '');
    },
    success: (message, data = null) => {
      const timestamp = new Date().toISOString();
      console.log(`[SUCCESS ${timestamp}] ${message}`, data || '');
    }
  };

  const getAndValidateInput = (key, message, defaultValue, advancedMode) => {
    if (advancedMode !== true) {
      return defaultValue;
    }
    const result = prompt(message, defaultValue);

    if (!result) {
      logger.warn('User cancelled input dialog or provided null value');
      alert(`No value for ${key} provided. Operation cancelled.`);
      return null;
    }

    if (result.trim() === '') {
      logger.warn('User provided empty string after trimming');
      alert(`Value for ${key} cannot be empty. Operation cancelled.`);
      return null;
    }

    logger.info(`Value for ${key} collected`, { string: result, length: result.length });
    return result;
  };

  let maxPages;

  // Function to get user input and confirmation with enhanced validation
  const getUserInput = () => {
    logger.info('Starting user input collection');

    try {
      maxPages = getAndValidateInput('maxPages', 'Max calendar pages (months/weeks) to process', 12, true);
      const trimmedSearch = getAndValidateInput(
        'trimmedSearch',
        'Enter the text/string to search for in calendar events that you want to delete',
        '',
        true
      );

      const confirmation = confirm(
        `Are you sure you want to delete ALL events containing "${trimmedSearch}"?\n\n` +
          'This action cannot be undone. The script will:\n' +
          `1. Search through up to ${maxPages} pages of your calendar\n` +
          '2. Delete every event that contains this text\n' +
          '3. Continue until all matching events are removed\n\n' +
          'Click OK to proceed or Cancel to abort.'
      );

      if (!confirmation) {
        logger.info('User declined confirmation dialog');
        return null;
      }

      logger.success('User input validated and confirmed', { searchString: trimmedSearch });
      return trimmedSearch;
    } catch (error) {
      logger.error('Error during user input collection', error);
      alert('An error occurred while collecting input. Please try again.');
      return null;
    }
  };

  // Enhanced element selection with retry logic
  const waitForElement = (selector, timeout = 5000, retryInterval = 100) => {
    return new Promise((resolve, reject) => {
      const startTime = Date.now();

      const checkElement = () => {
        try {
          const element = document.querySelector(selector);
          if (element) {
            logger.info('Element found', { selector, timeWaited: Date.now() - startTime });
            resolve(element);
            return;
          }

          if (Date.now() - startTime >= timeout) {
            logger.warn('Element not found within timeout', { selector, timeout });
            reject(new Error(`Element ${selector} not found within ${timeout}ms`));
            return;
          }

          setTimeout(checkElement, retryInterval);
        } catch (error) {
          logger.error('Error while waiting for element', { selector, error });
          reject(error);
        }
      };

      checkElement();
    });
  };

  // Simulates pressing Escape, to close any leftover context menu or
  // dialog before starting the next action. This helps avoid the
  // browser falling back to its own native right-click menu when a
  // stray Outlook menu from the previous action is still around.
  const simulateEscape = () => {
    const opts = { bubbles: true, cancelable: true, key: 'Escape', code: 'Escape', keyCode: 27, which: 27 };
    document.dispatchEvent(new KeyboardEvent('keydown', opts));
    document.dispatchEvent(new KeyboardEvent('keyup', opts));
  };

  // Simulates a right-click / secondary click (a two-finger click on a
  // Mac trackpad does the same thing) to open Outlook's context menu on
  // an event: Categorize, Private, Duplicate event, Save as .ics, Delete.
  //
  // IMPORTANT: we compute the event's actual on-screen position and pass
  // it as clientX/clientY. Without this, the simulated click defaults to
  // coordinate (0,0), which Outlook interprets as clicking near the
  // top-left of the page (around the ribbon/toolbar) — producing a
  // generic menu with the wrong options instead of the event's own menu.
  const simulateRightClick = (el) => {
    el.scrollIntoView({ block: 'center', inline: 'center' });
    const rect = el.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    const opts = {
      bubbles: true,
      cancelable: true,
      view: window,
      button: 2,
      clientX: x,
      clientY: y,
      screenX: x,
      screenY: y
    };
    el.dispatchEvent(new MouseEvent('mousedown', opts));
    el.dispatchEvent(new MouseEvent('mouseup', opts));
    el.dispatchEvent(new MouseEvent('contextmenu', opts));
  };

  // Recursively collects elements whose exact trimmed text matches
  // `text`, including elements nested inside any OPEN shadow DOM (some
  // modern Microsoft UI components render their internals inside a
  // shadow root, which normal DOM queries from the page can't see into
  // at all unless we explicitly walk into el.shadowRoot).
  const deepFindByExactText = (root, text, results = []) => {
    const children = root.querySelectorAll('*');
    for (const el of children) {
      if ((el.textContent || '').trim() === text) {
        results.push(el);
      }
      if (el.shadowRoot) {
        deepFindByExactText(el.shadowRoot, text, results);
      }
    }
    return results;
  };

  // Finds any clickable-looking element by its visible text (exact
  // match, trimmed). We no longer restrict this to <button> or
  // role="menuitem" — Outlook's menu items don't always use those, so
  // instead we search everything and pick the most specific (fewest
  // descendants) match, since a match with lots of children is usually
  // a wrapping container rather than the actual clickable row.
  const waitForButtonByText = (text, timeout = 5000, retryInterval = 100) => {
    return new Promise((resolve, reject) => {
      const startTime = Date.now();

      const checkButtons = () => {
        try {
          const candidates = deepFindByExactText(document, text);

          if (candidates.length > 0) {
            candidates.sort(
              (a, b) => a.querySelectorAll('*').length - b.querySelectorAll('*').length
            );
            const match = candidates[0];
            logger.info('Element found by text', {
              text,
              tag: match.tagName,
              timeWaited: Date.now() - startTime
            });
            resolve(match);
            return;
          }

          if (Date.now() - startTime >= timeout) {
            logger.warn('Element not found by text within timeout', { text, timeout });
            reject(new Error(`Element with text "${text}" not found within ${timeout}ms`));
            return;
          }

          setTimeout(checkButtons, retryInterval);
        } catch (error) {
          logger.error('Error while waiting for element by text', { text, error });
          reject(error);
        }
      };

      checkButtons();
    });
  };

  // Enhanced XPath search with error handling.
  // Outlook Web renders event titles inside elements with role="button"
  // and an aria-label containing the subject/time/location, so we search
  // both visible text spans AND aria-label attributes to catch events
  // whose title isn't in a plain text node.
  const findMatchingEvents = (searchString) => {
    try {
      logger.info('Searching for matching events', { searchString });

      const xpath =
        `//*[@role='button' and (contains(@aria-label, '${searchString}') ` +
        `or contains(text(), '${searchString}'))]`;

      const result = document.evaluate(
        xpath,
        document,
        null,
        XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,
        null
      );

      logger.info('XPath search completed', {
        xpath,
        matchCount: result.snapshotLength,
        searchString
      });

      return result;
    } catch (error) {
      logger.error('Error during XPath search', { searchString, error });
      throw new Error(`Failed to search for events: ${error.message}`);
    }
  };

  // Enhanced event deletion with comprehensive error handling
  const deleteEvent = async (eventElement, eventIndex, totalFound) => {
    try {
      const eventText = eventElement.getAttribute('aria-label') || eventElement.textContent || 'Unknown event';
      logger.info('Attempting to delete event', {
        eventIndex: eventIndex + 1,
        totalFound,
        eventText: eventText.substring(0, 100) + (eventText.length > 100 ? '...' : '')
      });

      // Close any leftover menu/dialog from a previous action before
      // starting, so this right-click lands cleanly on the event.
      simulateEscape();
      await new Promise((resolve) => setTimeout(resolve, 200));

      // Right-click (or two-finger click on a trackpad) the event to open
      // its context menu (Categorize, Private, Duplicate event, Save as
      // .ics, Delete, etc.). If the menu doesn't show up correctly the
      // first time (e.g. the browser's own native menu appears instead
      // of Outlook's), press Escape and try once more before giving up.
      let menuDeleteButton;
      try {
        simulateRightClick(eventElement);
        logger.info('Event right-clicked, waiting for context menu Delete option');
        menuDeleteButton = await waitForButtonByText(SELECTORS.deleteEventButtonText, 3000);
      } catch (firstAttemptError) {
        logger.warn('Context menu Delete option not found on first try, retrying once', {
          error: firstAttemptError.message
        });
        simulateEscape();
        await new Promise((resolve) => setTimeout(resolve, 400));
        simulateRightClick(eventElement);
        logger.info('Event right-clicked again, waiting for context menu Delete option');
        menuDeleteButton = await waitForButtonByText(SELECTORS.deleteEventButtonText, 3000);
      }
      menuDeleteButton.click();
      logger.info('Context menu Delete clicked, waiting for confirmation dialog');

      // Give the menu a moment to close and the confirmation dialog
      // ("Delete event" / "Are you sure you want to delete this event?")
      // a moment to render, so we don't accidentally re-click the
      // menu item that's mid-way through disappearing.
      await new Promise((resolve) => setTimeout(resolve, 400));

      // Click "Delete" again — this time in the confirmation dialog
      const confirmDeleteButton = await waitForButtonByText(SELECTORS.deleteEventButtonText, 3000);
      confirmDeleteButton.click();
      logger.info('Confirmation Delete clicked, checking for recurring event dialog');

      // Wait and check for recurring event dialog
      return new Promise((resolve) => {
        setTimeout(async () => {
          try {
            const recurringDialog = document.querySelector(SELECTORS.recurringDialogSelector);

            const dialogIndicatesRecurring =
              recurringDialog &&
              SELECTORS.recurringDialogIndicatorText.some((text) =>
                recurringDialog.textContent.toLowerCase().includes(text.toLowerCase())
              );

            if (dialogIndicatesRecurring) {
              logger.info('Recurring event dialog detected');

              const confirmButton = await waitForButtonByText(
                SELECTORS.recurringDialogConfirmButtonText,
                2000
              );
              confirmButton.click();

              logger.success('Recurring event deleted', { eventText });
              resolve(true);
            } else {
              logger.success('Regular event deleted', { eventText });
              resolve(true);
            }
          } catch (dialogError) {
            logger.error('Error handling deletion dialog', { eventText, error: dialogError });
            // Still consider it a success if we got this far
            resolve(true);
          }
        }, 500);
      });
    } catch (error) {
      logger.error('Error deleting event', {
        eventIndex: eventIndex + 1,
        eventText: eventElement?.getAttribute('aria-label') || eventElement?.textContent || 'Unknown',
        error
      });
      throw error;
    }
  };

  // Enhanced page navigation with error handling
  const navigateToNextPage = async (currentPage) => {
    try {
      logger.info('Navigating to next page', { currentPage, nextPage: currentPage + 1 });

      const nextButton = await waitForElement(SELECTORS.nextPageButton, 3000);
      nextButton.click();

      logger.info('Next button clicked, waiting for page load');
      await new Promise((resolve) => setTimeout(resolve, 2000));

      logger.success('Navigation completed', { newPage: currentPage + 1 });
      return true;
    } catch (error) {
      logger.error('Error navigating to next page', { currentPage, error });
      return false;
    }
  };

  // Main deletion process with comprehensive error handling
  const processCalendarDeletion = async (searchString) => {
    let currentPage = 1;
    let totalDeleted = 0;
    let totalErrors = 0;
    const errors = [];

    logger.info('Starting calendar deletion process', {
      searchString,
      maxPages,
      startTime: new Date().toISOString()
    });

    const processPage = async () => {
      try {
        if (currentPage > maxPages) {
          logger.success('All pages processed', {
            totalPages: maxPages,
            totalDeleted,
            totalErrors,
            completionTime: new Date().toISOString()
          });

          let message = `Deletion complete!\nTotal events deleted: ${totalDeleted}`;
          if (totalErrors > 0) {
            message += `\nErrors encountered: ${totalErrors} (check console for details)`;
          }
          alert(message);
          return;
        }

        logger.info('Processing page', { currentPage, maxPages });

        const matchingSpans = findMatchingEvents(searchString);
        const matchCount = matchingSpans.snapshotLength;

        if (matchCount > 0) {
          logger.info('Found matching events on page', { currentPage, matchCount });

          try {
            const firstMatch = matchingSpans.snapshotItem(0);
            await deleteEvent(firstMatch, 0, matchCount);
            totalDeleted++;

            logger.success('Event deleted successfully', {
              currentPage,
              totalDeleted,
              remainingOnPage: matchCount - 1
            });

            // Continue processing the same page after a short delay
            setTimeout(() => processPage(), 1500);
          } catch (deleteError) {
            totalErrors++;
            errors.push({
              page: currentPage,
              error: deleteError.message,
              timestamp: new Date().toISOString()
            });

            logger.error('Failed to delete event, continuing', {
              currentPage,
              totalErrors,
              error: deleteError
            });

            setTimeout(() => processPage(), 1500);
          }
        } else {
          logger.info('No matching events found on page', { currentPage });

          if (currentPage < maxPages) {
            const navigationSuccess = await navigateToNextPage(currentPage);

            if (navigationSuccess) {
              currentPage++;
              setTimeout(() => processPage(), 2000);
            } else {
              logger.error('Failed to navigate to next page, ending process', { currentPage });
              alert(`Process stopped due to navigation error on page ${currentPage}.\nTotal events deleted: ${totalDeleted}`);
            }
          } else {
            logger.success('Reached maximum pages', { maxPages, totalDeleted, totalErrors });

            let message = `Process complete! Processed ${maxPages} pages.\nTotal events deleted: ${totalDeleted}`;
            if (totalErrors > 0) {
              message += `\nErrors encountered: ${totalErrors} (check console for details)`;
            }
            alert(message);
          }
        }
      } catch (pageError) {
        totalErrors++;
        errors.push({
          page: currentPage,
          error: pageError.message,
          timestamp: new Date().toISOString()
        });

        logger.error('Error processing page', { currentPage, error: pageError });

        if (currentPage < maxPages) {
          logger.warn('Attempting to continue with next page after error');
          try {
            const navigationSuccess = await navigateToNextPage(currentPage);
            if (navigationSuccess) {
              currentPage++;
              setTimeout(() => processPage(), 3000);
            } else {
              logger.error('Cannot continue - navigation failed');
              alert(`Process stopped due to critical error on page ${currentPage}.\nTotal events deleted: ${totalDeleted}\nCheck console for error details.`);
            }
          } catch (navError) {
            logger.error('Critical error - cannot continue', navError);
            alert(`Critical error occurred. Process stopped.\nTotal events deleted: ${totalDeleted}\nCheck console for details.`);
          }
        } else {
          logger.error('Error on final page, ending process', { totalDeleted, totalErrors });
          alert(`Process completed with errors.\nTotal events deleted: ${totalDeleted}\nErrors: ${totalErrors}\nCheck console for details.`);
        }
      }
    };

    await processPage();
  };

  // Main execution with top-level error handling
  const main = async () => {
    try {
      logger.info('Outlook calendar deletion script started');

      const isOutlookHost =
        window.location.hostname.includes('outlook.office.com') ||
        window.location.hostname.includes('outlook.office365.com') ||
        window.location.hostname.includes('outlook.live.com');

      if (!isOutlookHost) {
        const warning = 'This script is designed for Outlook Web Calendar. Current page may not be supported.';
        logger.warn(warning);
        if (!confirm(warning + '\n\nDo you want to continue anyway?')) {
          logger.info('User chose not to continue on non-Outlook page');
          return;
        }
      }

      const searchString = getUserInput();
      if (!searchString) {
        logger.info('Operation cancelled by user during input phase');
        return;
      }

      await processCalendarDeletion(searchString);
    } catch (error) {
      logger.error('Critical error in main execution', error);
      alert('A critical error occurred. Check the browser console for details.');
    }
  };

  main();
})();
