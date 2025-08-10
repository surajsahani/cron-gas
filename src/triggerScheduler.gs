/**
 * TriggerScheduler Library
 * Create a hybrid scheduler for Google Apps Script:
 * - peak hours run at short interval
 * - off-peak run at longer (hourly) interval
 *
 * Usage:
 * TriggerScheduler.createHybridTrigger({
 *   func: 'scheduledPush', // the wrapper will call your target via property
 *   targetFunc: 'pushDataToBackend',
 *   peakStart: 6,
 *   peakEnd: 23,
 *   peakIntervalMinutes: 10,
 *   offPeakIntervalMinutes: 60,
 * });
 *
 * Then call setupTriggers() to create.
 */

var TriggerScheduler = (function () {

  function _propKey(targetFunc) {
    return 'TriggerScheduler_' + targetFunc;
  }

  function createHybridTrigger(options) {
    if (!options || !options.targetFunc) throw new Error('options.targetFunc is required');

    // Save options
    PropertiesService.getScriptProperties().setProperty(_propKey(options.targetFunc), JSON.stringify(options));

    // Remove existing wrapper triggers for this target
    deleteTriggersForWrapper(options.targetFunc);

    // determine base interval for creating the time-based trigger
    // we create a trigger with the smallest interval we need so wrapper can gate runs
    var baseInterval = Math.min(options.peakIntervalMinutes || 10, options.offPeakIntervalMinutes || 60);
    if (baseInterval < 1) baseInterval = 1;

    // wrapper name
    var wrapper = wrapperName(options.targetFunc);

    // ensure wrapper function exists in global scope
    _ensureWrapper(options.targetFunc);

    // create the trigger calling wrapper every `baseInterval` minutes
    ScriptApp.newTrigger(wrapper)
      .timeBased()
      .everyMinutes(baseInterval)
      .create();
  }

  function deleteHybridTrigger(targetFunc) {
    PropertiesService.getScriptProperties().deleteProperty(_propKey(targetFunc));
    deleteTriggersForWrapper(targetFunc);
  }

  function deleteTriggersForWrapper(targetFunc) {
    var wrapper = wrapperName(targetFunc);
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function (t) {
      try {
        if (t.getHandlerFunction && t.getHandlerFunction() === wrapper) {
          ScriptApp.deleteTrigger(t);
        }
      } catch (e) {}
    });
  }

  function wrapperName(targetFunc) {
    return 'TriggerSchedulerWrapper_' + targetFunc;
  }

  function _ensureWrapper(targetFunc) {
    var name = wrapperName(targetFunc);
    if (typeof this[name] === 'function') return;

    var self = this;
    // create global wrapper that reads schedule from PropertiesService and decides whether to call target
    this[name] = function () {
      var prop = PropertiesService.getScriptProperties().getProperty(_propKey(targetFunc));
      if (!prop) return;
      var options = JSON.parse(prop);

      var now = new Date();
      var hour = now.getHours();     // 0-23
      var minute = now.getMinutes(); // 0-59
      var isPeak = (hour >= (options.peakStart || 6) && hour < (options.peakEnd || 23));

      if (isPeak) {
        // run during peak windows
        if (typeof self[options.targetFunc] === 'function') {
          try { self[options.targetFunc](); } catch (e) { console.error(e); }
        } else {
          // If target function not present, try to run by name via global eval
          try { this[options.targetFunc] && this[options.targetFunc](); } catch (e) {}
        }
      } else {
        // off-peak: only run at specified offPeakIntervalMinutes
        var off = options.offPeakIntervalMinutes || 60;
        // run if current minute aligns to off-peak interval (for example minute === 0 for hourly)
        if (minute % off === 0) {
          if (typeof self[options.targetFunc] === 'function') {
            try { self[options.targetFunc](); } catch (e) { console.error(e); }
          } else {
            try { this[options.targetFunc] && this[options.targetFunc](); } catch (e) {}
          }
        }
      }
    };
  }

  return {
    createHybridTrigger: createHybridTrigger,
    deleteHybridTrigger: deleteHybridTrigger
  };
})();
