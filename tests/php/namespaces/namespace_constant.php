<?php
// vybe-test: php/namespaces/namespace_constant
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Config;
const VERSION = '1.0.0';
const MAX_RETRIES = 3;
echo VERSION . ' retries=' . MAX_RETRIES;
