# vybe-test: python/py_logging/test_py_logging_dict_config
# origin: languages/python/tests/python/test_py_logging.rs

import logging, logging.config, io

buf = io.StringIO()

config = {
    "version": 1,
    "disable_existing_loggers": False,
    "formatters": {
        "simple": {"format": "%(levelname)s - %(message)s"}
    },
    "handlers": {
        "console": {
            "class": "logging.StreamHandler",
            "formatter": "simple",
            "stream": "ext://sys.stdout"
        }
    },
    "loggers": {
        "myapp": {"level": "INFO", "handlers": ["console"], "propagate": False}
    }
}

logging.config.dictConfig(config)
logger = logging.getLogger("myapp")
# Just verify it's configured correctly
print(logger.level == logging.INFO)
print(len(logger.handlers) >= 1)
