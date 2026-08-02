<?php
// vybe-test: php/phase2/attribute_on_class
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

#[Entity] #[Table('users')] class User {}
