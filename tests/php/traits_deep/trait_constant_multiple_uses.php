<?php
// vybe-test: php/traits_deep/trait_constant_multiple_uses
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait StatusCodes {
    const OK    = 200;
    const ERROR = 500;
}
class ApiA { use StatusCodes; }
class ApiB { use StatusCodes; }
echo ApiA::OK . ',' . ApiB::ERROR;
