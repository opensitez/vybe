//! PHP 8+ multi-type catch unions, ordering, interfaces, Error vs Exception.

crate::php_cases! {
    catch_union_type_error_or_value_error => {
        r#"<?php
function risky(bool $flag): void {
    if ($flag) { throw new TypeError('type'); }
    throw new ValueError('value');
}
foreach ([true, false] as $f) {
    try { risky($f); }
    catch (TypeError | ValueError $e) { echo $e->getMessage() . ','; }
}
"#,
        ["type,value,"]
    };

    catch_order_specific_before_general_exception => {
        r#"<?php
try { throw new RuntimeException('rt'); }
catch (RuntimeException $e) { echo 'specific'; }
catch (Exception $e) { echo 'general'; }
"#,
        ["specific"]
    };

    catch_order_general_never_reached_if_specific_matches => {
        r#"<?php
try { throw new LogicException('logic'); }
catch (Exception $e) { echo 'general'; }
catch (LogicException $e) { echo 'specific'; }
"#,
        ["general"]
    };

    catch_union_runtime_or_logic_exception => {
        r#"<?php
function boom(int $n): void {
    if ($n === 1) { throw new RuntimeException('run'); }
    throw new LogicException('logic');
}
foreach ([1, 2] as $n) {
    try { boom($n); }
    catch (RuntimeException | LogicException $e) { echo $e->getMessage(); }
}
"#,
        ["runlogic"]
    };

    catch_interface_when_class_implements_it => {
        r#"<?php
interface Retryable {}
class HttpFailure extends Exception implements Retryable {}
try { throw new HttpFailure('503'); }
catch (Retryable $r) { echo 'retryable'; }
"#,
        ["retryable"]
    };

    catch_concrete_class_not_interface_alone_on_unrelated => {
        r#"<?php
interface Marker {}
class Plain extends Exception {}
try { throw new Plain('plain'); }
catch (Marker $m) { echo 'marker'; }
catch (Exception $e) { echo 'exception'; }
"#,
        ["exception"]
    };

    catch_error_not_exception_for_type_error => {
        r#"<?php
try { throw new TypeError('te'); }
catch (Exception $e) { echo 'exception'; }
catch (Error $e) { echo 'error'; }
"#,
        ["error"]
    };

    catch_exception_only_misses_error_subclass => {
        r#"<?php
$log = [];
try {
    throw new DivisionByZeroError('div0');
} catch (Exception $e) {
    $log[] = 'exception';
} catch (Error $e) {
    $log[] = 'error';
}
echo implode(',', $log);
"#,
        ["error"]
    };

    catch_throwable_catches_exception => {
        r#"<?php
try { throw new Exception('ex'); }
catch (Throwable $t) { echo get_class($t); }
"#,
        ["Exception"]
    };

    catch_throwable_catches_error => {
        r#"<?php
try { throw new ParseError('parse'); }
catch (Throwable $t) { echo get_class($t); }
"#,
        ["ParseError"]
    };

    catch_throwable_after_specific_skipped_when_unmatched => {
        r#"<?php
try { throw new ValueError('ve'); }
catch (TypeError $e) { echo 'type'; }
catch (Throwable $t) { echo 'all'; }
"#,
        ["all"]
    };

    catch_exception_when_throwing_error_subclass_uncaught => {
        r#"<?php
$log = [];
try {
    try { throw new ArithmeticError('arith'); }
    catch (Exception $e) { $log[] = 'caught exc'; }
} catch (Error $e) {
    $log[] = 'caught err';
}
echo implode(',', $log);
"#,
        ["caught err"]
    };

    catch_union_parent_or_child_matches_child => {
        r#"<?php
class BaseEx extends Exception {}
class DerivedEx extends BaseEx {}
try { throw new DerivedEx('d'); }
catch (BaseEx | DerivedEx $e) { echo 'union hit'; }
"#,
        ["union hit"]
    };

    catch_union_with_unrelated_second_type => {
        r#"<?php
try { throw new OverflowException('of'); }
catch (UnderflowException | OverflowException $e) { echo $e->getMessage(); }
"#,
        ["of"]
    };

    catch_parse_error_or_type_error_union => {
        r#"<?php
function fail(int $mode): void {
    if ($mode === 1) { throw new ParseError('parse'); }
    throw new TypeError('type');
}
foreach ([1, 2] as $m) {
    try { fail($m); }
    catch (ParseError | TypeError $e) { echo $e->getMessage(); }
}
"#,
        ["parsetype"]
    };

    catch_invalid_argument_or_domain_union => {
        r#"<?php
function validate(int $v): void {
    if ($v < 0) { throw new InvalidArgumentException('arg'); }
    throw new DomainException('domain');
}
try { validate(-1); } catch (InvalidArgumentException | DomainException $e) { echo 'a'; }
try { validate(1); } catch (InvalidArgumentException | DomainException $e) { echo 'd'; }
"#,
        ["ad"]
    };

    catch_error_or_exception_union_on_exception => {
        r#"<?php
try { throw new RuntimeException('rt'); }
catch (Error | Exception $e) { echo 'either'; }
"#,
        ["either"]
    };

    catch_error_or_exception_union_on_error => {
        r#"<?php
try { throw new AssertionError('assert'); }
catch (Error | Exception $e) { echo 'either'; }
"#,
        ["either"]
    };

    catch_throwable_alone_as_catch_all => {
        r#"<?php
function probe(bool $useError): void {
    if ($useError) { throw new Error('e'); }
    throw new Exception('x');
}
try { probe(true); } catch (Throwable $t) { echo 'E'; }
try { probe(false); } catch (Throwable $t) { echo 'X'; }
"#,
        ["EX"]
    };

    multiple_union_catch_blocks_in_order => {
        r#"<?php
function step(int $n): void {
    if ($n === 1) { throw new TypeError('t'); }
    if ($n === 2) { throw new ValueError('v'); }
    throw new RuntimeException('r');
}
foreach ([1, 2, 3] as $n) {
    try { step($n); }
    catch (TypeError | ValueError $e) { echo 'tv'; }
    catch (RuntimeException $e) { echo 'rt'; }
}
"#,
        ["tvtvrt"]
    };

    catch_division_by_zero_in_arithmetic_union => {
        r#"<?php
try { throw new DivisionByZeroError('zero'); }
catch (ArithmeticError | DivisionByZeroError $e) { echo 'arith'; }
"#,
        ["arith"]
    };

    catch_interface_implemented_by_custom_exception => {
        r#"<?php
interface Loggable { public function logMessage(): string; }
class AuditEvent extends Exception implements Loggable {
    public function logMessage(): string { return 'audit'; }
}
try { throw new AuditEvent('evt'); }
catch (Loggable $l) { echo $l->logMessage(); }
"#,
        ["audit"]
    };

    catch_wrong_union_types_miss_and_fall_through => {
        r#"<?php
try { throw new LogicException('logic'); }
catch (RuntimeException | InvalidArgumentException $e) { echo 'miss'; }
catch (LogicException $e) { echo 'hit'; }
"#,
        ["hit"]
    };

    catch_union_first_matching_type_wins => {
        r#"<?php
try { throw new RuntimeException('r'); }
catch (RuntimeException | Exception $e) { echo get_class($e); }
"#,
        ["RuntimeException"]
    };

    catch_custom_classes_in_union => {
        r#"<?php
class Alpha extends Exception {}
class Beta extends Exception {}
try { throw new Beta('b'); }
catch (Alpha | Beta $e) { echo $e->getMessage(); }
"#,
        ["b"]
    };

    catch_union_then_separate_single_type_block => {
        r#"<?php
try { throw new OutOfRangeException('oor'); }
catch (InvalidArgumentException | DomainException $e) { echo 'union'; }
catch (OutOfRangeException $e) { echo 'single'; }
"#,
        ["single"]
    };

    error_subclass_not_caught_by_exception_only_handler => {
        r#"<?php
$handled = false;
try {
    throw new CompileError('compile');
} catch (Exception $e) {
    $handled = true;
}
echo $handled ? 'handled' : 'unhandled';
"#,
        ["unhandled"]
    };

    error_caught_when_union_includes_error => {
        r#"<?php
try { throw new CompileError('compile'); }
catch (Exception | Error $e) { echo 'caught'; }
"#,
        ["caught"]
    };

    exception_caught_by_throwable_handler => {
        r#"<?php
try { throw new BadFunctionCallException('bad fn'); }
catch (Throwable $t) { echo $t->getMessage(); }
"#,
        ["bad fn"]
    };

    catch_specific_error_before_throwable => {
        r#"<?php
try { throw new TypeError('type'); }
catch (TypeError $e) { echo 'specific'; }
catch (Throwable $t) { echo 'general'; }
"#,
        ["specific"]
    };

    catch_throwable_after_exception_block_on_error => {
        r#"<?php
try { throw new Error('plain error'); }
catch (Exception $e) { echo 'exc'; }
catch (Throwable $t) { echo 'thr'; }
"#,
        ["thr"]
    };

    catch_union_with_interface_and_class => {
        r#"<?php
interface Stoppable {}
class StopEvent extends Exception implements Stoppable {}
try { throw new StopEvent('stop'); }
catch (Stoppable | RuntimeException $e) { echo 'stopped'; }
"#,
        ["stopped"]
    };

    catch_runtime_or_logic_on_logic_throw => {
        r#"<?php
try { throw new LogicException('logic'); }
catch (RuntimeException | LogicException $e) { echo 'logic'; }
"#,
        ["logic"]
    };

    catch_runtime_or_logic_on_runtime_throw => {
        r#"<?php
try { throw new RuntimeException('runtime'); }
catch (RuntimeException | LogicException $e) { echo 'runtime'; }
"#,
        ["runtime"]
    };

    catch_value_error_or_type_error_on_value => {
        r#"<?php
try { throw new ValueError('bad value'); }
catch (ValueError | TypeError $e) { echo 'value'; }
"#,
        ["value"]
    };

    catch_value_error_or_type_error_on_type => {
        r#"<?php
try { throw new TypeError('bad type'); }
catch (ValueError | TypeError $e) { echo 'type'; }
"#,
        ["type"]
    };

    catch_exception_subclass_via_parent_union_member => {
        r#"<?php
class ServiceFailure extends RuntimeException {}
try { throw new ServiceFailure('svc'); }
catch (RuntimeException | LogicException $e) { echo 'svc'; }
"#,
        ["svc"]
    };

    catch_only_exception_when_throwing_error_reaches_outer => {
        r#"<?php
$log = [];
try {
    try { throw new TypeError('inner'); }
    catch (Exception $e) { $log[] = 'inner exc'; }
} catch (TypeError $e) {
    $log[] = 'outer type';
}
echo implode(',', $log);
"#,
        ["outer type"]
    };

    catch_error_subclass_with_error_union => {
        r#"<?php
try { throw new DivisionByZeroError('div'); }
catch (DivisionByZeroError | ArithmeticError $e) { echo 'div'; }
"#,
        ["div"]
    };

    catch_exception_chain_with_union_and_fallback => {
        r#"<?php
function work(int $step): void {
    if ($step === 1) { throw new InvalidArgumentException('a'); }
    if ($step === 2) { throw new UnexpectedValueException('u'); }
    throw new RuntimeException('r');
}
foreach ([1, 2, 3] as $s) {
    try { work($s); }
    catch (InvalidArgumentException | UnexpectedValueException $e) { echo 'iu'; }
    catch (RuntimeException $e) { echo 'rt'; }
}
"#,
        ["iurt"]
    };

    catch_interface_on_non_exception_fails_next_handler => {
        r#"<?php
interface NotThrowable {}
try { throw new RuntimeException('rt'); }
catch (NotThrowable $n) { echo 'iface'; }
catch (RuntimeException $e) { echo 'rt'; }
"#,
        ["rt"]
    };

    catch_throwable_variable_accessible_in_handler => {
        r#"<?php
try { throw new Exception('msg', 42); }
catch (Throwable $t) { echo $t->getMessage() . ':' . $t->getCode(); }
"#,
        ["msg:42"]
    };

    catch_union_without_variable => {
        r#"<?php
try { throw new ValueError('v'); }
catch (TypeError | ValueError) { echo 'union no var'; }
"#,
        ["union no var"]
    };

    catch_multiple_unions_exclusive_handlers => {
        r#"<?php
function mode(int $m): void {
    if ($m === 1) { throw new TypeError('t'); }
    if ($m === 2) { throw new ValueError('v'); }
    throw new RuntimeException('r');
}
foreach ([1, 2, 3] as $m) {
    try { mode($m); }
    catch (TypeError | ValueError $e) { echo 'A'; }
    catch (RuntimeException | LogicException $e) { echo 'B'; }
}
"#,
        ["AABB"]
    };

    catch_error_before_exception_on_throwable_error => {
        r#"<?php
try { throw new Error('err'); }
catch (Error $e) { echo 'error first'; }
catch (Exception $e) { echo 'exception'; }
"#,
        ["error first"]
    };

    catch_exception_before_error_on_exception => {
        r#"<?php
try { throw new Exception('ex'); }
catch (Exception $e) { echo 'exception first'; }
catch (Error $e) { echo 'error'; }
"#,
        ["exception first"]
    };

    catch_underflow_or_overflow_on_overflow => {
        r#"<?php
try { throw new OverflowException('high'); }
catch (UnderflowException | OverflowException $e) { echo 'overflow'; }
"#,
        ["overflow"]
    };

    catch_underflow_or_overflow_on_underflow => {
        r#"<?php
try { throw new UnderflowException('low'); }
catch (UnderflowException | OverflowException $e) { echo 'underflow'; }
"#,
        ["underflow"]
    };

    catch_bad_method_call_via_exception_union => {
        r#"<?php
class Demo {}
try { (new Demo())->missing(); }
catch (BadMethodCallException | Error $e) { echo get_class($e); }
"#,
        ["Error"]
    };

    catch_throwable_on_builtin_value_error => {
        r#"<?php
try {
    $arr = [1];
    echo $arr[5];
} catch (Throwable $t) {
    echo 'caught';
}
"#,
        ["caught"]
    };

    catch_union_parent_interface_and_child_class => {
        r#"<?php
interface ProblemDomain {}
class BillingProblem extends DomainException implements ProblemDomain {}
try { throw new BillingProblem('bill'); }
catch (ProblemDomain | DomainException $e) { echo 'billing'; }
"#,
        ["billing"]
    };

    catch_specific_parse_before_generic_throwable => {
        r#"<?php
try { throw new ParseError('syntax'); }
catch (ParseError $e) { echo 'parse'; }
catch (Throwable $t) { echo 'other'; }
"#,
        ["parse"]
    };

    catch_exception_when_error_subclass_thrown_stays_uncaught => {
        r#"<?php
$hit = '';
try { throw new AssertionError('fail'); }
catch (Exception $e) { $hit = 'exc'; }
catch (Error $e) { $hit = 'err'; }
echo $hit;
"#,
        ["err"]
    };

    catch_union_three_types_any_match => {
        r#"<?php
function tri(int $n): void {
    if ($n === 1) { throw new TypeError('t'); }
    if ($n === 2) { throw new ValueError('v'); }
    throw new ParseError('p');
}
foreach ([1, 2, 3] as $n) {
    try { tri($n); }
    catch (TypeError | ValueError | ParseError $e) { echo $n; }
}
"#,
        ["123"]
    };

    catch_runtime_exception_before_union_on_runtime => {
        r#"<?php
try { throw new RuntimeException('r'); }
catch (RuntimeException $e) { echo 'solo'; }
catch (RuntimeException | LogicException $e) { echo 'union'; }
"#,
        ["solo"]
    };

    catch_logic_exception_union_after_runtime_miss => {
        r#"<?php
try { throw new LogicException('l'); }
catch (RuntimeException $e) { echo 'rt'; }
catch (RuntimeException | LogicException $e) { echo 'union'; }
"#,
        ["union"]
    };

    catch_error_interface_not_valid_use_exception => {
        r#"<?php
try { throw new RangeException('range'); }
catch (Exception $e) { echo 'range ok'; }
"#,
        ["range ok"]
    };

    catch_throwable_last_resort_after_specific_miss => {
        r#"<?php
try { throw new JsonException('json'); }
catch (TypeError $e) { echo 'type'; }
catch (ValueError $e) { echo 'value'; }
catch (Throwable $t) { echo 'json'; }
"#,
        ["json"]
    };

    catch_union_on_extended_spl_exception => {
        r#"<?php
try { throw new UnexpectedValueException('uv'); }
catch (UnexpectedValueException | OutOfBoundsException $e) { echo 'uv'; }
"#,
        ["uv"]
    };

    catch_error_union_when_throwing_exception_only => {
        r#"<?php
try { throw new LengthException('len'); }
catch (Error | LengthException $e) { echo 'len'; }
"#,
        ["len"]
    };
}
