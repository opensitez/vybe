# vybe-test: python/filesystem_utils/configparser_raw
# origin: languages/python/tests/python/test_filesystem_utils.rs
# vybe-test-mode: compile

import configparser
c = configparser.ConfigParser()
c.read_string('[s]\nk=%(name)s\n', source='s')
