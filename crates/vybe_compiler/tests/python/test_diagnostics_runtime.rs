//! logging, warnings, traceback runtime diagnostics.

crate::runtime_case!(
    logging_getlogger,
    "import logging\nprint(logging.getLogger('test').name)\n",
    "test"
);
crate::runtime_case!(
    logging_level_info,
    "import logging\nprint(logging.INFO)\n",
    "20"
);
crate::runtime_case!(
    logging_level_debug,
    "import logging\nprint(logging.DEBUG < logging.INFO)\n",
    "True"
);
crate::runtime_case!(
    logging_basicconfig,
    "import logging\nlogging.basicConfig(level=logging.INFO)\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    logging_logger_warning,
    "import logging\nlog = logging.getLogger('t')\nprint(callable(log.warning))\n",
    "True"
);
crate::runtime_case!(
    logging_logger_error,
    "import logging\nlog = logging.getLogger('t')\nprint(callable(log.error))\n",
    "True"
);
crate::runtime_case!(
    logging_logger_debug,
    "import logging\nlog = logging.getLogger('t')\nprint(callable(log.debug))\n",
    "True"
);
crate::runtime_case!(
    logging_logger_info,
    "import logging\nlog = logging.getLogger('t')\nprint(callable(log.info))\n",
    "True"
);
crate::runtime_case!(
    logging_logger_critical,
    "import logging\nlog = logging.getLogger('t')\nprint(callable(log.critical))\n",
    "True"
);
crate::runtime_case!(
    logging_logrecord,
    "import logging\nprint(hasattr(logging, 'LogRecord'))\n",
    "True"
);
crate::runtime_case!(
    logging_formatter,
    "import logging\nprint(hasattr(logging, 'Formatter'))\n",
    "True"
);
crate::runtime_case!(
    logging_streamhandler,
    "import logging\nprint(hasattr(logging, 'StreamHandler'))\n",
    "True"
);
crate::runtime_case!(
    logging_filehandler,
    "import logging\nprint(hasattr(logging, 'FileHandler'))\n",
    "True"
);
crate::runtime_case!(
    logging_filter,
    "import logging\nprint(hasattr(logging, 'Filter'))\n",
    "True"
);
crate::runtime_case!(
    warnings_warn,
    "import warnings\nwith warnings.catch_warnings(record=True) as w:\n warnings.simplefilter('always')\n warnings.warn('msg')\n print(len(w))\n",
    "1"
);
crate::runtime_case!(
    warnings_warn_explicit,
    "import warnings\nprint(callable(warnings.warn_explicit))\n",
    "True"
);
crate::runtime_case!(
    warnings_filterwarnings,
    "import warnings\nwarnings.filterwarnings('ignore')\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    warnings_resetwarnings,
    "import warnings\nwarnings.resetwarnings()\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    warnings_simplefilter,
    "import warnings\nwarnings.simplefilter('default')\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    warnings_catch_warnings,
    "import warnings\nprint(hasattr(warnings, 'catch_warnings'))\n",
    "True"
);
crate::runtime_case!(
    warnings_warn_category,
    "import warnings\nwith warnings.catch_warnings(record=True) as w:\n warnings.simplefilter('always')\n warnings.warn('x', UserWarning)\n print(issubclass(w[0].category, UserWarning))\n",
    "True"
);
crate::runtime_case!(
    traceback_format_exc,
    "import traceback\ntry:\n 1/0\nexcept:\n s = traceback.format_exc()\n print('ZeroDivisionError' in s)\n",
    "True"
);
crate::runtime_case!(
    traceback_format_exception,
    "import traceback\ntry:\n raise ValueError('e')\nexcept:\n import sys\n lines = traceback.format_exception(*sys.exc_info())\n print(len(lines) > 0)\n",
    "True"
);
crate::runtime_case!(
    traceback_extract_tb,
    "import traceback\ntry:\n 1/0\nexcept:\n import sys\n print(hasattr(traceback, 'extract_tb'))\n",
    "True"
);
crate::runtime_case!(
    traceback_print_exc_capture,
    "import traceback\nimport io\ntry:\n 1/0\nexcept:\n buf = io.StringIO()\n traceback.print_exc(file=buf)\n print(len(buf.getvalue()) > 0)\n",
    "True"
);
crate::runtime_case!(
    traceback_format_tb,
    "import traceback\ntry:\n 1/0\nexcept:\n import sys\n tb = sys.exc_info()[2]\n print(len(traceback.format_tb(tb)) > 0)\n",
    "True"
);
crate::runtime_case!(
    traceback_format_stack,
    "import traceback\nprint(len(traceback.format_stack()) > 0)\n",
    "True"
);
crate::runtime_case!(
    traceback_limit,
    "import traceback\ntry:\n 1/0\nexcept:\n s = traceback.format_exc(limit=1)\n print(isinstance(s, str))\n",
    "True"
);
crate::runtime_case!(
    logging_root_logger,
    "import logging\nprint(logging.root.name)\n",
    "root"
);
crate::runtime_case!(
    logging_logger_level,
    "import logging\nlog = logging.getLogger('x')\nlog.setLevel(logging.ERROR)\nprint(log.level)\n",
    "40"
);
crate::runtime_case!(
    logging_logger_has_handlers,
    "import logging\nlog = logging.getLogger('y')\nprint(isinstance(log.handlers, list))\n",
    "True"
);
crate::runtime_case!(
    logging_last_resort,
    "import logging\nprint(hasattr(logging, 'lastResort'))\n",
    "True"
);
crate::runtime_case!(
    warnings_showwarning,
    "import warnings\nprint(callable(warnings.showwarning))\n",
    "True"
);
crate::runtime_case!(
    warnings_onceregistry,
    "import warnings\nprint(hasattr(warnings, '_filters_mutated'))\n",
    "True"
);
crate::runtime_case!(
    traceback_exception_class,
    "import traceback\nprint(hasattr(traceback, 'TracebackException'))\n",
    "True"
);
crate::runtime_case!(
    traceback_walk_tb,
    "import traceback\ntry:\n 1/0\nexcept:\n import sys\n print(hasattr(traceback, 'walk_tb'))\n",
    "True"
);
crate::runtime_case!(
    logging_exception_method,
    "import logging\nlog = logging.getLogger('z')\nprint(callable(log.exception))\n",
    "True"
);
crate::runtime_case!(
    logging_log_method,
    "import logging\nlog = logging.getLogger('w')\nprint(callable(log.log))\n",
    "True"
);
crate::runtime_case!(
    logging_is_enabled_for,
    "import logging\nlog = logging.getLogger('v')\nprint(callable(log.isEnabledFor))\n",
    "True"
);
crate::runtime_case!(
    warnings_deprecation,
    "import warnings\nprint(issubclass(DeprecationWarning, Warning))\n",
    "True"
);
crate::runtime_case!(
    warnings_runtime_warning,
    "import warnings\nprint(issubclass(RuntimeWarning, Warning))\n",
    "True"
);
crate::runtime_case!(
    traceback_clear_frames,
    "import traceback\ntry:\n 1/0\nexcept:\n import sys\n tb = sys.exc_info()[2]\n print(tb is not None)\n",
    "True"
);
crate::runtime_case!(
    logging_module_name,
    "import logging\nprint(logging.__name__)\n",
    "logging"
);
crate::runtime_case!(
    warnings_module_name,
    "import warnings\nprint(warnings.__name__)\n",
    "warnings"
);
crate::runtime_case!(
    traceback_module_name,
    "import traceback\nprint(traceback.__name__)\n",
    "traceback"
);
crate::runtime_case!(
    logging_add_level_name,
    "import logging\nprint(callable(logging.addLevelName))\n",
    "True"
);
crate::runtime_case!(
    logging_get_level_name,
    "import logging\nprint(logging.getLevelName(logging.INFO))\n",
    "INFO"
);

crate::compile_case!(
    logging_dictconfig,
    "import logging.config\nlogging.config.dictConfig({})\n"
);
crate::compile_case!(
    logging_handlers_rotating,
    "import logging.handlers\nlogging.handlers.RotatingFileHandler\n"
);
crate::compile_case!(
    traceback_print_stack,
    "import traceback\ntraceback.print_stack()\n"
);
crate::compile_case!(
    warnings_warn_stacklevel,
    "import warnings\nwarnings.warn('m', stacklevel=2)\n"
);
crate::compile_case!(
    logging_logrecord_factory,
    "import logging\nlogging.setLogRecordFactory(logging.LogRecord)\n"
);
