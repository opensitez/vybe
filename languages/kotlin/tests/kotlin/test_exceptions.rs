use crate::helpers::run_prints;

#[test]
fn test_try_finally_flow() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                println("try start")
            } finally {
                println("finally block")
            }
        }
    "#,
    );
    assert_eq!(out, &["try start", "finally block"]);
}

#[test]
fn test_try_catch_flow() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                println("try start")
                throw Exception("failure")
            } catch (e: Exception) {
                println("catch block")
            }
        }
    "#,
    );
    assert_eq!(out, &["try start", "catch block"]);
}

#[test]
fn test_try_catch_finally_combined() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                println("working")
                throw Exception("bad arg")
            } catch (e: Exception) {
                println("handled arg error")
            } finally {
                println("cleanup done")
            }
        }
    "#,
    );
    assert_eq!(out, &["working", "handled arg error", "cleanup done"]);
}

#[test]
fn test_try_without_exception() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                println("normal flow")
            } catch (e: Exception) {
                println("error")
            }
        }
    "#,
    );
    assert_eq!(out, &["normal flow"]);
}

#[test]
fn test_throw_and_catch_specific() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                throw Exception("boom")
            } catch (e: IllegalArgumentException) {
                println("arg")
            } catch (e: Exception) {
                println("general")
            } finally {
                println("complete")
            }
        }
    "#,
    );
    assert_eq!(out, &["general", "complete"]);
}

#[test]
fn test_require_helper() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                require(false)
            } catch (e: Exception) {
                println("require failed")
            }
        }
    "#,
    );
    assert_eq!(out, &["require failed"]);
}

#[test]
fn test_check_helper() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                check(false)
            } catch (e: Exception) {
                println("check failed")
            }
        }
    "#,
    );
    assert_eq!(out, &["check failed"]);
}

#[test]
fn test_nested_try_finally() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                try {
                    throw Exception("inner")
                } catch (e: Exception) {
                    println("inner catch")
                } finally {
                    println("inner finally")
                }
            } finally {
                println("outer finally")
            }
        }
    "#,
    );
    assert_eq!(out, &["inner catch", "inner finally", "outer finally"]);
}

#[test]
fn test_nested_try_catch_with_success_path() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                try {
                    println("inner")
                } catch (e: Exception) {
                    println("should not")
                } finally {
                    println("inner done")
                }
            } catch (e: Exception) {
                println("outer catch")
            } finally {
                println("outer done")
            }
        }
    "#,
    );
    assert_eq!(out, &["inner", "inner done", "outer done"]);
}

#[test]
fn test_try_catch_multiple_handlers() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                throw IllegalArgumentException("bad")
            } catch (e: IllegalArgumentException) {
                println("arg")
            } catch (e: Exception) {
                println("general")
            }
        }
    "#,
    );
    assert_eq!(out, &["arg"]);
}

#[test]
fn test_try_with_finally_no_exception() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                println("ok")
            } finally {
                println("cleanup")
            }
        }
    "#,
    );
    assert_eq!(out, &["ok", "cleanup"]);
}

#[test]
fn test_throw_in_nested_function() {
    let out = run_prints(
        r#"
        fun fail() {
            throw Exception("inner")
        }

        fun main() {
            try {
                fail()
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#,
    );
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_try_finally_only_with_error() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                throw Exception("panic")
            } finally {
                println("ended")
            }
        }
    "#,
    );
    assert_eq!(out, &["ended"]);
}

#[test]
fn test_multiple_exceptions_cascading() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                try {
                    throw IllegalArgumentException("bad")
                } catch (e: IllegalArgumentException) {
                    throw Exception("wrapped")
                }
            } catch (e: Exception) {
                println("wrapped")
            }
        }
    "#,
    );
    assert_eq!(out, &["wrapped"]);
}

#[test]
fn test_require_false_path_with_message() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                require(false)
            } catch (e: Exception) {
                println("require")
            }
            println("done")
        }
    "#,
    );
    assert_eq!(out, &["require", "done"]);
}

#[test]
fn test_check_false_path_in_nested_context() {
    let out = run_prints(
        r#"
        fun assertPositive(n: Int) {
            check(n > 0)
        }

        fun main() {
            try {
                assertPositive(0)
            } catch (e: Exception) {
                println("invalid")
            }
        }
    "#,
    );
    assert_eq!(out, &["invalid"]);
}

#[test]
fn test_try_catch_finally_with_return_value() {
    let out = run_prints(
        r#"
        fun calc(): Int {
            try {
                return 5
            } catch (e: Exception) {
                return 0
            } finally {
                println("finally")
            }
        }

        fun main() {
            println(calc())
        }
    "#,
    );
    assert_eq!(out, &["finally", "5"]);
}

#[test]
fn test_exception_no_throw_in_try() {
    let out = run_prints(
        r#"
fun main() {
    try {
        println("safe")
    } catch (e: Exception) {
        println("caught")
    } finally {
        println("done")
    }
}
"#,
    );
    assert_eq!(out, &["safe", "done"]);
}

#[test]
fn test_exception_catch_supertype() {
    let out = run_prints(
        r#"
fun main() {
    try {
        throw IllegalArgumentException("bad")
    } catch (e: Exception) {
        println("caught")
    }
}
"#,
    );
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_exception_with_else_path() {
    let out = run_prints(
        r#"
fun main() {
    try {
        throw Exception()
    } catch (e: IllegalArgumentException) {
        println("arg")
    } catch (e: Exception) {
        println("general")
    }
}
"#,
    );
    assert_eq!(out, &["general"]);
}

#[test]
fn test_exception_nested_finally_order() {
    let out = run_prints(
        r#"
fun main() {
    try {
        try {
            println("inner")
        } finally {
            println("inner finally")
        }
    } finally {
        println("outer finally")
    }
}
"#,
    );
    assert_eq!(out, &["inner", "inner finally", "outer finally"]);
}

#[test]
fn test_exception_try_in_function() {
    let out = run_prints(
        r#"
fun risky(value: Int) {
    if (value < 0) {
        throw Exception("no")
    }
}

fun main() {
    try {
        risky(-1)
    } catch (e: Exception) {
        println("blocked")
    }
}
"#,
    );
    assert_eq!(out, &["blocked"]);
}

#[test]
fn test_exception_return_value_ignored_by_finally() {
    let out = run_prints(
        r#"
fun result(): Int {
    try {
        return 10
    } finally {
        println("end")
    }
}

fun main() {
    println(result())
}
"#,
    );
    assert_eq!(out, &["end", "10"]);
}

#[test]
fn test_exception_multiple_guarded_calls() {
    let out = run_prints(
        r#"
fun fail(v: Int) {
    if (v == 1) {
        throw Exception("x")
    }
}

fun main() {
    for (v in arrayOf(0, 1, 2)) {
        try {
            fail(v)
            println(v)
        } catch (e: Exception) {
            println("err")
        }
    }
}
"#,
    );
    assert_eq!(out, &["0", "err", "2"]);
}

#[test]
fn test_exception_nested_catch_chained() {
    let out = run_prints(
        r#"
fun main() {
    try {
        throw Exception("top")
    } catch (e: Exception) {
        try {
            throw IllegalArgumentException("inner")
        } catch (e: IllegalArgumentException) {
            println("nested")
        }
    }
}
"#,
    );
    assert_eq!(out, &["nested"]);
}

#[test]
fn test_exception_finally_with_continue_path() {
    let out = run_prints(
        r#"
fun main() {
    for (v in 1..3) {
        try {
            println(v)
        } finally {
            println("tick")
        }
    }
}
"#,
    );
    assert_eq!(out, &["1", "tick", "2", "tick", "3", "tick"]);
}

#[test]
fn test_exception_try_finalize_no_catch() {
    let out = run_prints(
        r#"
fun main() {
    try {
        println("go")
    } finally {
        println("final")
    }
}
"#,
    );
    assert_eq!(out, &["go", "final"]);
}

#[test]
fn test_exception_require_like_guard() {
    let out = run_prints(
        r#"
fun main() {
    try {
        require(false)
    } catch (e: Exception) {
        println("guarded")
    } finally {
        println("released")
    }
}
"#,
    );
    assert_eq!(out, &["guarded", "released"]);
}

#[test]
fn test_exception_check_like_guard() {
    let out = run_prints(
        r#"
fun validate(ok: Boolean): Int {
    return if (ok) 1 else 0
}

fun main() {
    try {
        if (validate(false) == 0) {
            throw Exception("bad")
        }
    } catch (e: Exception) {
        println("bad")
    }
}
"#,
    );
    assert_eq!(out, &["bad"]);
}

#[test]
fn test_exception_nested_function_throwing() {
    let out = run_prints(
        r#"
fun inner(): Int {
    throw Exception("inner")
}

fun outer() {
    try {
        inner()
    } catch (e: Exception) {
        println("outer")
    }
}

fun main() {
    outer()
}
"#,
    );
    assert_eq!(out, &["outer"]);
}

#[test]
fn test_exception_resource_like_flow() {
    let out = run_prints(
        r#"
fun run() {
    try {
        println("open")
    } finally {
        println("closed")
    }
}

fun main() {
    run()
}
"#,
    );
    assert_eq!(out, &["open", "closed"]);
}

#[test]
fn test_exception_custom_exception_class_matches_catch() {
    let out = run_prints(
        r#"
class NetworkError(message: String) : Exception(message)

fun main() {
    try {
        throw NetworkError("down")
    } catch (e: NetworkError) {
        println("custom")
        println(e.message)
    }
}
"#,
    );
    assert_eq!(out, &["custom", "down"]);
}

#[test]
fn test_exception_constructor_failure_is_caught() {
    let out = run_prints(
        r#"
class Exploding {
    init {
        println("init")
        throw Exception("explode")
    }
}

fun main() {
    try {
        val _ = Exploding()
        println("constructed")
    } catch (e: Exception) {
        println("caught")
        println(e.message)
    }
}
"#,
    );
    assert_eq!(out, &["init", "caught", "explode"]);
}

#[test]
fn test_exception_throw_in_catch_and_finally_executes() {
    let out = run_prints(
        r#"
fun main() {
    try {
        try {
            throw Exception("inner")
        } catch (e: Exception) {
            println("caught-inner")
            throw Exception("from-catch")
        } finally {
            println("inner-finally")
        }
    } catch (e: Exception) {
        println("caught-outer")
        println(e.message)
    }
}
"#,
    );
    assert_eq!(
        out,
        &[
            "caught-inner",
            "inner-finally",
            "caught-outer",
            "from-catch"
        ]
    );
}

#[test]
fn test_exception_throw_in_finally_overrides_body_exception() {
    let out = run_prints(
        r#"
fun main() {
    try {
        try {
            throw Exception("body")
        } finally {
            println("body-finally")
            throw Exception("finally")
        }
    } catch (e: Exception) {
        println(e.message)
    }
}
"#,
    );
    assert_eq!(out, &["body-finally", "finally"]);
}

#[test]
fn test_exception_finally_with_loop_break() {
    let out = run_prints(
        r#"
fun main() {
    for (value in 1..4) {
        try {
            println(value)
            if (value == 3) {
                break
            }
        } finally {
            println("finally")
        }
    }
    println("done")
}
"#,
    );
    assert_eq!(
        out,
        &["1", "finally", "2", "finally", "3", "finally", "done"]
    );
}

#[test]
fn test_exception_finally_does_not_modify_returned_value_binding() {
    let out = run_prints(
        r#"
fun compute(): Int {
    var result = 1
    try {
        println("try")
        result = 5
        return result
    } finally {
        println("finally")
        result = 9
    }
}

fun main() {
    println(compute())
}
"#,
    );
    assert_eq!(out, &["try", "finally", "5"]);
}

#[test]
fn test_exception_throw_in_outer_catch_is_represented() {
    let out = run_prints(
        r#"
fun main() {
    try {
        try {
            throw Exception("inner")
        } catch (e: Exception) {
            throw Exception("outer")
        }
    } catch (e: Exception) {
        println("caught")
        println(e.message)
    }
}
"#,
    );
    assert_eq!(out, &["caught", "outer"]);
}

#[test]
fn test_exception_finally_around_return_from_nested_context() {
    let out = run_prints(
        r#"
fun evaluate(): Int {
    return try {
        throw Exception("primary")
    } catch (e: Exception) {
        println("inner-catch")
        7
    } finally {
        println("inner-finally")
    }
}

fun main() {
    println(evaluate())
}
"#,
    );
    assert_eq!(out, &["inner-catch", "inner-finally", "7"]);
}

#[test]
fn test_exception_nested_finally_and_catch_cleanup_order() {
    let out = run_prints(
        r#"
fun main() {
    try {
        try {
            println("inner-try")
            throw Exception("inner")
        } catch (e: Exception) {
            println("inner-catch")
            throw e
        } finally {
            println("inner-finally")
        }
    } catch (e: Exception) {
        println("outer-catch")
    } finally {
        println("outer-finally")
    }
}
"#,
    );
    assert_eq!(
        out,
        &[
            "inner-try",
            "inner-catch",
            "inner-finally",
            "outer-catch",
            "outer-finally"
        ]
    );
}

#[test]
fn test_exception_catch_variable_scope_isolated_from_outer() {
    let out = run_prints(
        r#"
fun main() {
    val message = "root"
    try {
        throw Exception("inner")
    } catch (message: Exception) {
        println(message.message)
    }
    println("root")
}
"#,
    );
    assert_eq!(out, &["inner", "root"]);
}

#[test]
fn test_exception_try_with_continue_in_catch_then_finally() {
    let out = run_prints(
        r#"
fun main() {
    for (value in 1..4) {
        try {
            if (value == 2) {
                throw Exception("bad")
            }
            println(value)
            continue
        } catch (e: Exception) {
            println("caught")
            continue
        } finally {
            println("finally")
        }
    }
    println("done")
}
"#,
    );
    assert_eq!(
        out,
        &[
            "1", "finally", "caught", "finally", "3", "finally", "4", "finally", "done"
        ]
    );
}

#[test]
fn test_exception_require_and_finally_cleanup() {
    let out = run_prints(
        r#"
fun main() {
    try {
        try {
            throw Exception("boom")
        } finally {
            println("cleanup")
        }
    } catch (e: Exception) {
        println("caught")
    }
}
    "#,
    );
    assert_eq!(out, &["cleanup", "caught"]);
}

#[test]
fn test_run_catching_recover_with_default() {
    let out = run_prints(
        r#"
        fun main() {
            val bad = runCatching {
                throw Exception("oops")
            }
            println(bad.isSuccess)
            println(bad.isFailure)
            println(bad.getOrElse { "fallback" })
        }
    "#,
    );
    assert_eq!(out, &["false", "true", "fallback"]);
}

#[test]
fn test_run_catching_ok_path_preserves_value() {
    let out = run_prints(
        r#"
        fun main() {
            val good = runCatching {
                val value = 3 + 4
                value * 2
            }
            println(good.isSuccess)
            println(good.getOrNull())
            println(good.getOrElse { 0 })
        }
    "#,
    );
    assert_eq!(out, &["true", "14", "14"]);
}

#[test]
fn test_try_expression_as_value_with_finally_cleanup() {
    let out = run_prints(
        r#"
        fun main() {
            val value = try {
                println("body")
                9
            } finally {
                println("closed")
            }
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["body", "closed", "9"]);
}

#[test]
fn test_try_expression_value_on_failure_path() {
    let out = run_prints(
        r#"
        fun main() {
            val value = try {
                throw Exception("boom")
            } catch (e: Exception) {
                5
            }
            println(value)
            println("done")
        }
    "#,
    );
    assert_eq!(out, &["5", "done"]);
}

#[test]
fn test_try_expression_finally_with_no_catch() {
    let out = run_prints(
        r#"
        fun status(flag: Boolean): String {
            return try {
                if (flag) "ok" else throw Exception("bad")
            } finally {
                println("cleanup")
            }
        }

        fun main() {
            try {
                println(status(false))
            } catch (e: Exception) {
                println(e.message)
            }
        }
    "#,
    );
    assert_eq!(out, &["cleanup", "bad"]);
}

#[test]
fn test_run_catching_map_and_recover_flow() {
    let out = run_prints(
        r#"
        fun main() {
            val value = runCatching { "k".toInt() }
                .map { it + 1 }
                .onFailure { println("fail") }
                .recover { 9 }

            println(value.getOrNull())
            println(value.isFailure)
        }
    "#,
    );
    assert_eq!(out, &["fail", "9", "false"]);
}

#[test]
fn test_run_catching_with_typed_match() {
    let out = run_prints(
        r#"
        fun main() {
            val result = runCatching {
                throw IllegalArgumentException("bad")
            }

            val message = result.exceptionOrNull()?.let { it.message } ?: "none"
            println(result.isFailure)
            println(message)
        }
    "#,
    );
    assert_eq!(out, &["true", "bad"]);
}

#[test]
fn test_exception_type_hierarchy_catch_order() {
    let out = run_prints(
        r#"
        class BaseError : Exception("base")
        class DerivedError : BaseError()

        fun main() {
            try {
                throw DerivedError()
            } catch (e: DerivedError) {
                println("derived")
            } catch (e: BaseError) {
                println("base")
            } catch (e: Exception) {
                println("general")
            }
        }
    "#,
    );
    assert_eq!(out, &["derived"]);
}
