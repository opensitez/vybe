<?php
// vybe-test: php/exception_types/multiple_catch_different_types
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

function riskyOp(int $kind): void {
    if ($kind === 1) throw new InvalidArgumentException('bad arg');
    if ($kind === 2) throw new RuntimeException('runtime');
    throw new LogicException('logic');
}
foreach ([1, 2, 3] as $k) {
    try {
        riskyOp($k);
    } catch (InvalidArgumentException $e) {
        echo 'invalid';
    } catch (RuntimeException $e) {
        echo 'runtime';
    } catch (LogicException $e) {
        echo 'logic';
    }
}
