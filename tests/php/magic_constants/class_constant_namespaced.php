<?php
// vybe-test: php/magic_constants/class_constant_namespaced
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

namespace Http;
class Request {}
echo Request::class; // Http\Request
