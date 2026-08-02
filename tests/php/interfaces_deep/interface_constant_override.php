<?php
// vybe-test: php/interfaces_deep/interface_constant_override
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface HasDefault { const string MODE = 'default'; }
class Custom implements HasDefault { const string MODE = 'custom'; }
class Default_ implements HasDefault {}
echo Custom::MODE . ':' . Default_::MODE;
