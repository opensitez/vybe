<?php
// vybe-test: php/classes/class_final_class_compilation_only_runtime
// origin: languages/php/tests/php/test_classes.rs
// vybe-test-mode: compile

final class Sealed {}
new Sealed();
