# vybe-test: python/filesystem_utils/configparser_raw
# origin: languages/python/tests/python/test_filesystem_utils.rs

import configparser
c = configparser.ConfigParser()
c.read_string('[s]\nk=%(name)s\n', source='s')
