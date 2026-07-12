//! Deeply nested try/catch handlers, rethrow, transforms, and catch/finally interplay.

crate::php_cases! {
    nested_two_levels_inner_catch_swallows => {
        r#"<?php
$log = [];
try {
    try { throw new Exception('inner'); }
    catch (Exception $e) { $log[] = 'inner'; }
} catch (Exception $e) {
    $log[] = 'outer';
}
echo implode(',', $log);
"#,
        ["inner"]
    };

    nested_two_levels_inner_rethrows_outer_catches => {
        r#"<?php
$log = [];
try {
    try { throw new Exception('msg'); }
    catch (Exception $e) { $log[] = 'inner'; throw $e; }
} catch (Exception $e) {
    $log[] = 'outer:' . $e->getMessage();
}
echo implode(',', $log);
"#,
        ["inner,outer:msg"]
    };

    nested_three_levels_deepest_catch_only => {
        r#"<?php
$log = [];
try {
    try {
        try { throw new RuntimeException('deep'); }
        catch (RuntimeException $e) { $log[] = 'L3'; }
    } catch (RuntimeException $e) {
        $log[] = 'L2';
    }
} catch (RuntimeException $e) {
    $log[] = 'L1';
}
echo implode(',', $log);
"#,
        ["L3"]
    };

    nested_three_levels_middle_rethrows_to_outer => {
        r#"<?php
$log = [];
try {
    try {
        try { throw new Exception('x'); }
        catch (Exception $e) { $log[] = 'in'; throw $e; }
    } catch (Exception $e) {
        $log[] = 'mid';
        throw $e;
    }
} catch (Exception $e) {
    $log[] = 'out';
}
echo implode(',', $log);
"#,
        ["in,mid,out"]
    };

    nested_four_levels_all_rethrow_chain => {
        r#"<?php
$log = [];
try {
    try {
        try {
            try { throw new LogicException('4'); }
            catch (LogicException $e) { $log[] = 'c4'; throw $e; }
        } catch (LogicException $e) { $log[] = 'c3'; throw $e; }
    } catch (LogicException $e) { $log[] = 'c2'; throw $e; }
} catch (LogicException $e) { $log[] = 'c1'; }
echo implode(',', $log);
"#,
        ["c4,c3,c2,c1"]
    };

    catch_transforms_exception_message => {
        r#"<?php
try {
    try { throw new RuntimeException('raw'); }
    catch (RuntimeException $e) {
        throw new Exception('wrapped:' . $e->getMessage());
    }
} catch (Exception $e) {
    echo $e->getMessage();
}
"#,
        ["wrapped:raw"]
    };

    catch_transforms_exception_preserves_code => {
        r#"<?php
try {
    try { throw new RuntimeException('r', 404); }
    catch (RuntimeException $e) {
        throw new Exception('mapped', $e->getCode());
    }
} catch (Exception $e) {
    echo $e->getMessage() . ':' . $e->getCode();
}
"#,
        ["mapped:404"]
    };

    inner_catch_does_not_catch_outer_throw => {
        r#"<?php
$log = [];
try {
    try { echo 'inner,'; }
    catch (Exception $e) { $log[] = 'inner catch'; }
    throw new Exception('outer throw');
} catch (Exception $e) {
    $log[] = 'outer catch';
}
echo implode(',', $log);
"#,
        ["inner,outer catch"]
    };

    try_without_catch_with_finally_propagates => {
        r#"<?php
$log = [];
try {
    try { throw new Exception('up'); }
    finally { $log[] = 'finally'; }
} catch (Exception $e) {
    $log[] = 'caught';
}
echo implode(',', $log);
"#,
        ["finally,caught"]
    };

    multiple_catch_blocks_first_matching_runtime => {
        r#"<?php
try { throw new RuntimeException('rt'); }
catch (LogicException $e) { echo 'logic'; }
catch (RuntimeException $e) { echo 'runtime'; }
catch (Exception $e) { echo 'generic'; }
"#,
        ["runtime"]
    };

    multiple_catch_blocks_falls_through_to_generic => {
        r#"<?php
try { throw new Exception('gen'); }
catch (RuntimeException $e) { echo 'runtime'; }
catch (LogicException $e) { echo 'logic'; }
catch (Exception $e) { echo 'generic'; }
"#,
        ["generic"]
    };

    catch_variable_used_in_finally_block => {
        r#"<?php
$msg = '';
try {
    throw new Exception('stored');
} catch (Exception $e) {
    $msg = $e->getMessage();
} finally {
    echo 'finally:' . $msg;
}
"#,
        ["finally:stored"]
    };

    catch_variable_scope_not_visible_after_block => {
        r#"<?php
try { throw new Exception('scoped'); }
catch (Exception $e) { $inside = $e->getMessage(); }
finally { echo isset($inside) ? $inside : 'gone'; }
"#,
        ["scoped"]
    };

    nested_try_inside_catch_block => {
        r#"<?php
$log = [];
try {
    throw new Exception('first');
} catch (Exception $e) {
    $log[] = 'c1';
    try {
        throw new RuntimeException('second');
    } catch (RuntimeException $r) {
        $log[] = 'c2';
    }
}
echo implode(',', $log);
"#,
        ["c1,c2"]
    };

    sibling_try_blocks_independent_handlers => {
        r#"<?php
$log = [];
try { throw new Exception('a'); } catch (Exception $e) { $log[] = 'A'; }
try { throw new RuntimeException('b'); } catch (RuntimeException $e) { $log[] = 'B'; }
echo implode(',', $log);
"#,
        ["A,B"]
    };

    inner_throw_type_mismatch_reaches_outer_catch => {
        r#"<?php
try {
    try { throw new InvalidArgumentException('iae'); }
    catch (RuntimeException $e) { echo 'wrong'; }
} catch (InvalidArgumentException $e) {
    echo $e->getMessage();
}
"#,
        ["iae"]
    };

    middle_catch_rethrows_to_outer_handler => {
        r#"<?php
$log = [];
try {
    try { throw new DomainException('d'); }
    catch (DomainException $e) { $log[] = 'mid'; throw $e; }
} catch (DomainException $e) {
    $log[] = 'top';
}
echo implode(',', $log);
"#,
        ["mid,top"]
    };

    inner_finally_runs_before_inner_catch_on_throw => {
        r#"<?php
$log = [];
try {
    try { throw new Exception('e'); }
    finally { $log[] = 'fin'; }
} catch (Exception $ex) {
    $log[] = 'catch';
}
echo implode(',', $log);
"#,
        ["fin,catch"]
    };

    outer_catch_after_inner_try_finally_throw => {
        r#"<?php
$log = [];
try {
    try { $log[] = 'body'; }
    finally { throw new OverflowException('of'); }
} catch (OverflowException $e) {
    $log[] = 'outer';
}
echo implode(',', $log);
"#,
        ["body,outer"]
    };

    inner_catch_exception_only_misses_error => {
        r#"<?php
$log = [];
try {
    try { throw new TypeError('te'); }
    catch (Exception $e) { $log[] = 'exc'; }
} catch (TypeError $e) {
    $log[] = 'type';
}
echo implode(',', $log);
"#,
        ["type"]
    };

    nested_runtime_then_generic_catch_order => {
        r#"<?php
try {
    try { throw new RuntimeException('r'); }
    catch (RuntimeException $e) { echo 'inner rt'; }
} catch (Exception $e) {
    echo 'outer ex';
}
"#,
        ["inner rt"]
    };

    catch_block_throws_different_type_to_outer => {
        r#"<?php
try {
    try { throw new Exception('orig'); }
    catch (Exception $e) { throw new LogicException('new'); }
} catch (LogicException $e) {
    echo $e->getMessage();
}
"#,
        ["new"]
    };

    empty_inner_catch_swallows_silently => {
        r#"<?php
$log = [];
try {
    try { throw new Exception('gone'); }
    catch (Exception $e) {}
    $log[] = 'after inner';
} catch (Exception $e) {
    $log[] = 'outer';
}
echo implode(',', $log);
"#,
        ["after inner"]
    };

    catch_without_variable_in_nested_handler => {
        r#"<?php
$log = [];
try {
    try { throw new Exception('x'); }
    catch (Exception) { $log[] = 'no var'; }
} catch (Exception $e) {
    $log[] = 'outer';
}
echo implode(',', $log);
"#,
        ["no var"]
    };

    nested_handlers_with_custom_exception_hierarchy => {
        r#"<?php
class AppError extends Exception {}
class AppFatal extends AppError {}
try {
    try { throw new AppFatal('fatal'); }
    catch (AppError $e) { echo 'app:' . $e->getMessage(); }
} catch (Exception $e) {
    echo 'generic';
}
"#,
        ["app:fatal"]
    };

    closure_throw_from_nested_try_caught_outer => {
        r#"<?php
$fn = function () {
    try {
        try { throw new RuntimeException('closure'); }
        catch (RuntimeException $e) { throw $e; }
    } catch (RuntimeException $e) {
        echo $e->getMessage();
    }
};
$fn();
"#,
        ["closure"]
    };

    function_call_depth_with_nested_handlers => {
        r#"<?php
function inner() { throw new Exception('from inner'); }
function middle() {
    try { inner(); }
    catch (Exception $e) { throw new RuntimeException('middle:' . $e->getMessage()); }
}
try { middle(); }
catch (RuntimeException $e) { echo $e->getMessage(); }
"#,
        ["middle:from inner"]
    };

    loop_nested_try_each_iteration_catches => {
        r#"<?php
$log = [];
foreach ([1, 2, 3] as $n) {
    try {
        try { if ($n === 2) { throw new Exception('two'); } }
        catch (Exception $e) { $log[] = "c$n"; }
    } catch (Exception $e) {
        $log[] = "o$n";
    }
}
echo implode(',', $log);
"#,
        ["c2"]
    };

    switch_case_with_nested_try_catch => {
        r#"<?php
$log = [];
switch (1) {
    case 1:
        try {
            try { throw new Exception('sw'); }
            catch (Exception $e) { $log[] = 'inner'; }
        } catch (Exception $e) {
            $log[] = 'outer';
        }
}
echo implode(',', $log);
"#,
        ["inner"]
    };

    outer_catch_sees_exception_after_inner_finally_rethrow => {
        r#"<?php
$log = [];
try {
    try {
        throw new RuntimeException('inner');
    } finally {
        $log[] = 'fin';
        throw new LogicException('rethrow');
    }
} catch (LogicException $e) {
    $log[] = 'logic';
} catch (RuntimeException $e) {
    $log[] = 'runtime';
}
echo implode(',', $log);
"#,
        ["fin,logic"]
    };

    outer_try_inner_try_outer_catch_only_on_inner_throw => {
        r#"<?php
$log = [];
try {
    try { $log[] = 'run'; throw new Exception('x'); }
    catch (Exception $e) { $log[] = 'inner'; }
} catch (Exception $e) {
    $log[] = 'outer';
}
echo implode(',', $log);
"#,
        ["run,inner"]
    };

    three_separate_catch_blocks_same_try_ordered => {
        r#"<?php
function probe(int $kind): string {
    try {
        if ($kind === 1) { throw new InvalidArgumentException('a'); }
        if ($kind === 2) { throw new RuntimeException('b'); }
        throw new LogicException('c');
    } catch (InvalidArgumentException $e) { return 'invalid'; }
    catch (RuntimeException $e) { return 'runtime'; }
    catch (LogicException $e) { return 'logic'; }
}
echo probe(1) . probe(2) . probe(3);
"#,
        ["invalidruntimelogic"]
    };

    nested_handlers_in_instance_method => {
        r#"<?php
class Service {
    public function handle(): string {
        try {
            try { throw new Exception('svc'); }
            catch (Exception $e) { return 'inner'; }
        } catch (Exception $e) {
            return 'outer';
        }
    }
}
echo (new Service())->handle();
"#,
        ["inner"]
    };

    static_method_nested_catch_rethrow => {
        r#"<?php
class Api {
    public static function call(): void {
        try {
            try { throw new RuntimeException('api'); }
            catch (RuntimeException $e) { throw new Exception('mapped'); }
        } catch (Exception $e) {
            echo $e->getMessage();
        }
    }
}
Api::call();
"#,
        ["mapped"]
    };

    catch_variable_shadowing_in_nested_blocks => {
        r#"<?php
$log = [];
try {
    throw new Exception('outer');
} catch (Exception $e) {
    $log[] = $e->getMessage();
    try {
        throw new RuntimeException('inner');
    } catch (RuntimeException $e) {
        $log[] = $e->getMessage();
    }
}
echo implode(',', $log);
"#,
        ["outer,inner"]
    };

    nested_handler_preserves_previous_exception_chain => {
        r#"<?php
try {
    try { throw new RuntimeException('root', 7); }
    catch (RuntimeException $e) {
        throw new LogicException('wrap', 0, $e);
    }
} catch (LogicException $e) {
    echo $e->getPrevious()->getMessage() . ':' . $e->getPrevious()->getCode();
}
"#,
        ["root:7"]
    };

    inner_try_no_throw_outer_stays_quiet => {
        r#"<?php
$log = [];
try {
    try { $log[] = 'ok'; }
    catch (Exception $e) { $log[] = 'inner'; }
} catch (Exception $e) {
    $log[] = 'outer';
}
echo implode(',', $log);
"#,
        ["ok"]
    };

    inner_catch_logs_outer_catch_logs_on_rethrow => {
        r#"<?php
$log = [];
try {
    try { throw new Exception('up'); }
    catch (Exception $e) { $log[] = 'i'; throw $e; }
} catch (Exception $e) {
    $log[] = 'o';
}
echo implode(',', $log);
"#,
        ["i,o"]
    };

    finally_in_inner_catch_in_outer => {
        r#"<?php
$log = [];
try {
    try { throw new Exception('e'); }
    catch (Exception $ex) {
        try { $log[] = 'handled'; }
        finally { $log[] = 'fin'; }
    }
} catch (Exception $e) {
    $log[] = 'outer';
}
echo implode(',', $log);
"#,
        ["handled,fin"]
    };

    outer_throw_not_seen_by_already_finished_inner => {
        r#"<?php
$log = [];
try {
    try {
        $log[] = 'inner done';
    } catch (Exception $e) {
        $log[] = 'inner catch';
    }
    throw new Exception('after inner');
} catch (Exception $e) {
    $log[] = 'outer';
}
echo implode(',', $log);
"#,
        ["inner done,outer"]
    };

    nested_four_levels_inner_swallow_all_outer_quiet => {
        r#"<?php
$log = [];
try {
    try {
        try {
            try { throw new Exception('deep'); }
            catch (Exception $e) { $log[] = 'got'; }
        } catch (Exception $e) { $log[] = 'L3'; }
    } catch (Exception $e) { $log[] = 'L2'; }
} catch (Exception $e) { $log[] = 'L1'; }
echo implode(',', $log);
"#,
        ["got"]
    };

    catch_in_finally_wrapping_outer_failure => {
        r#"<?php
$log = [];
try {
    try { throw new Exception('try'); }
    finally {
        try { throw new RuntimeException('finally throw'); }
        catch (RuntimeException $e) { $log[] = 'fin catch'; }
    }
} catch (Exception $e) {
    $log[] = 'outer';
}
echo implode(',', $log);
"#,
        ["fin catch,outer"]
    };

    nested_try_with_different_exception_classes => {
        r#"<?php
class NetworkError extends Exception {}
class TimeoutError extends NetworkError {}
try {
    try { throw new TimeoutError('timeout'); }
    catch (NetworkError $e) { echo 'network'; }
} catch (Exception $e) {
    echo 'generic';
}
"#,
        ["network"]
    };

    inner_catch_runtime_outer_catch_exception_on_rethrow => {
        r#"<?php
try {
    try { throw new RuntimeException('rt'); }
    catch (RuntimeException $e) {
        throw new Exception('promoted');
    }
} catch (Exception $e) {
    echo $e->getMessage();
}
"#,
        ["promoted"]
    };

    nested_handlers_with_error_then_exception_layers => {
        r#"<?php
$log = [];
try {
    try { throw new ParseError('parse'); }
    catch (Exception $e) { $log[] = 'exc'; }
} catch (Error $e) {
    $log[] = 'err';
}
echo implode(',', $log);
"#,
        ["err"]
    };

    try_in_catch_rethrows_to_same_level_outer => {
        r#"<?php
$log = [];
try {
    throw new Exception('first');
} catch (Exception $e) {
    $log[] = 'c1';
    try {
        throw new RuntimeException('second');
    } catch (RuntimeException $r) {
        $log[] = 'c2';
        throw $r;
    }
} catch (RuntimeException $e) {
    $log[] = 'c3';
}
echo implode(',', $log);
"#,
        ["c1,c2"]
    };

    deeply_nested_no_throw_runs_all_finally => {
        r#"<?php
$log = [];
try {
    try {
        try { $log[] = 'body'; }
        finally { $log[] = 'f3'; }
    } finally { $log[] = 'f2'; }
} finally { $log[] = 'f1'; }
echo implode(',', $log);
"#,
        ["body,f3,f2,f1"]
    };

    inner_handler_transforms_then_outer_reads_message => {
        r#"<?php
try {
    try { throw new UnderflowException('low'); }
    catch (UnderflowException $e) {
        throw new OverflowException('high');
    }
} catch (OverflowException $e) {
    echo $e->getMessage();
}
"#,
        ["high"]
    };

    nested_optional_catch_outer_named_on_rethrow => {
        r#"<?php
$log = [];
try {
    try { throw new Exception('x'); }
    catch (Exception) { $log[] = 'inner'; throw new Exception('y'); }
} catch (Exception $e) {
    $log[] = $e->getMessage();
}
echo implode(',', $log);
"#,
        ["inner,y"]
    };

    four_level_mixed_catch_and_finally_order => {
        r#"<?php
$log = [];
try {
    try {
        try { throw new Exception('t'); }
        finally { $log[] = 'f3'; }
    } catch (Exception $e) { $log[] = 'c2'; throw $e; }
} catch (Exception $e) { $log[] = 'c1'; }
echo implode(',', $log);
"#,
        ["f3,c2,c1"]
    };

    catch_binds_used_in_nested_finally_after_rethrow => {
        r#"<?php
$saved = '';
try {
    try { throw new Exception('bind'); }
    catch (Exception $e) {
        $saved = $e->getMessage();
        throw $e;
    } finally {
        echo 'fin:' . $saved . ',';
    }
} catch (Exception $e) {
    echo 'out';
}
"#,
        ["fin:bind,out"]
    };

    inner_try_finally_no_throw_outer_catch_unused => {
        r#"<?php
$log = [];
try {
    try { $log[] = 'work'; }
    finally { $log[] = 'cleanup'; }
} catch (Exception $e) {
    $log[] = 'fail';
}
echo implode(',', $log);
"#,
        ["work,cleanup"]
    };

    nested_handler_with_interface_typed_catch => {
        r#"<?php
interface Problem {}
class ConcreteProblem extends Exception implements Problem {}
try {
    try { throw new ConcreteProblem('iface'); }
    catch (Problem $p) { echo 'iface ok'; }
} catch (Exception $e) {
    echo 'miss';
}
"#,
        ["iface ok"]
    };

    outer_catch_after_inner_multiple_catch_types => {
        r#"<?php
try {
    try { throw new ValueError('ve'); }
    catch (TypeError $e) { echo 'type'; }
    catch (ValueError $e) { echo 'value'; }
} catch (Exception $e) {
    echo 'exception';
}
"#,
        ["value"]
    };

    nested_rethrow_loses_inner_handler_after_finally => {
        r#"<?php
$log = [];
try {
    try {
        throw new Exception('a');
    } finally {
        $log[] = 'f';
    }
} catch (Exception $e) {
    $log[] = 'c';
}
echo implode(',', $log);
"#,
        ["f,c"]
    };

    transform_in_inner_adds_context_for_outer => {
        r#"<?php
try {
    try { throw new InvalidArgumentException('bad id'); }
    catch (InvalidArgumentException $e) {
        throw new DomainException('service:' . $e->getMessage());
    }
} catch (DomainException $e) {
    echo $e->getMessage();
}
"#,
        ["service:bad id"]
    };
}
