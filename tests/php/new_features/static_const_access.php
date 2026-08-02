<?php
// vybe-test: php/new_features/static_const_access
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

class Cfg { const VER = '1.0'; } echo Cfg::VER;
