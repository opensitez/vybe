use crate::helpers::run_prints;

#[test]
fn test_result_success_flags_and_get_or_null() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.success(7)
            println(value.isSuccess)
            println(value.isFailure)
            println(value.getOrNull())
            println(value.exceptionOrNull() == null)
        }
    "#);
    assert_eq!(out, &[
        "true",
        "false",
        "7",
        "true",
    ]);
}

#[test]
fn test_result_failure_flags_and_exception_or_null() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.failure<Int>(IllegalStateException("bad"))
            println(value.isSuccess)
            println(value.isFailure)
            println(value.getOrNull() == null)
            println(value.exceptionOrNull() is? Exception)
        }
    "#);
    assert_eq!(out, &[
        "false",
        "true",
        "true",
        "true",
    ]);
}

#[test]
fn test_result_get_or_else_failure_path() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.failure<Int>(Exception("boom"))
            println(value.getOrElse { 99 })
        }
    "#);
    assert_eq!(out, &["99"]);
}

#[test]
fn test_result_get_or_default_success_and_failure() {
    let out = run_prints(r#"
        fun main() {
            val success = Result.success(2)
            val fail = Result.failure<Int>(Exception("boom"))
            println(success.getOrDefault(9))
            println(fail.getOrDefault(9))
        }
    "#);
    assert_eq!(out, &["2", "9"]);
}

#[test]
fn test_result_map_transforms_success() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.success(3).map { it + 4 }
            println(value.getOrNull())
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_result_map_does_not_transform_failure() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.failure<Int>(Exception("boom")).map { it + 1 }
            println(value.isSuccess)
            println(value.exceptionOrNull()?.message)
        }
    "#);
    assert_eq!(out, &["false", "boom"]);
}

#[test]
fn test_result_map_catching_catches_mapper_exception() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.success(0).mapCatching { throw Exception("mapped") }
            println(value.isSuccess)
            println(value.exceptionOrNull()?.message)
        }
    "#);
    assert_eq!(out, &["false", "mapped"]);
}

#[test]
fn test_result_recover_overwrites_failure() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.failure<Int>(Exception("bad")).recover { 11 }
            println(value.isSuccess)
            println(value.getOrNull())
        }
    "#);
    assert_eq!(out, &["true", "11"]);
}

#[test]
fn test_result_recover_keeps_success() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.success(2).recover { 11 }
            println(value.isSuccess)
            println(value.getOrNull())
        }
    "#);
    assert_eq!(out, &["true", "2"]);
}

#[test]
fn test_result_recover_catching_with_success() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.failure<Int>(Exception("bad")).recoverCatching { throw Exception("oops") }
            println(value.isSuccess)
            println(value.exceptionOrNull()?.message)
        }
    "#);
    assert_eq!(out, &["false", "oops"]);
}

#[test]
fn test_result_on_success_side_effect() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.success("ok").onSuccess { println("hit") }
            value.onFailure { println("no") }
            println(value.getOrNull())
        }
    "#);
    assert_eq!(out, &["hit", "ok"]);
}

#[test]
fn test_result_on_failure_side_effect() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.failure<String>(Exception("bad")).onSuccess { println("hit") }
            value.onFailure { println("fail") }
            println(value.isFailure)
        }
    "#);
    assert_eq!(out, &["fail", "true"]);
}

#[test]
fn test_result_fold_success() {
    let out = run_prints(r#"
        fun main() {
            val value = runCatching { 5 }
            val output = value.fold({ "s:" + it.toString() }, { "f" })
            println(output)
        }
    "#);
    assert_eq!(out, &["s:5"]);
}

#[test]
fn test_result_fold_failure() {
    let out = run_prints(r#"
        fun main() {
            val value = runCatching { throw Exception("no") }
            val output = value.fold({ "s" + it.toString() }, { "f" })
            println(output)
        }
    "#);
    assert_eq!(out, &["f"]);
}

#[test]
fn test_result_get_or_throw_success() {
    let out = run_prints(r#"
        fun main() {
            val value = runCatching { 21 }
            println(value.getOrThrow())
        }
    "#);
    assert_eq!(out, &["21"]);
}

#[test]
fn test_result_get_or_throw_failure_throws() {
    let out = run_prints(r#"
        fun main() {
            val value = runCatching { throw IllegalArgumentException("bad") }
            try {
                value.getOrThrow()
                println("ok")
            } catch (e: IllegalArgumentException) {
                println(e.message)
            }
        }
    "#);
    assert_eq!(out, &["bad"]);
}

#[test]
fn test_result_failure_exception_message_is_preserved() {
    let out = run_prints(r#"
        fun main() {
            val value = runCatching { throw Exception("preserve") }
            println(value.exceptionOrNull()?.message)
        }
    "#);
    assert_eq!(out, &["preserve"]);
}

#[test]
fn test_result_from_companion_success_factory() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.success(99)
            println(value.getOrNull())
        }
    "#);
    assert_eq!(out, &["99"]);
}

#[test]
fn test_result_from_companion_failure_factory() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.failure<Int>(Exception("factory"))
            println(value.isFailure)
            println(value.exceptionOrNull()?.message)
        }
    "#);
    assert_eq!(out, &["true", "factory"]);
}

#[test]
fn test_result_with_run_catching_success_value() {
    let out = run_prints(r#"
        fun main() {
            val value = runCatching { "x" + "y" }
            println(value.getOrNull())
        }
    "#);
    assert_eq!(out, &["xy"]);
}

#[test]
fn test_run_catching_failure_path_message() {
    let out = run_prints(r#"
        fun main() {
            val value = runCatching { "1".toInt() + "a".toInt() }
            println(value.isFailure)
            println(value.exceptionOrNull()?.javaClass?.simpleName)
        }
    "#);
    assert_eq!(out, &["true", "NumberFormatException"]);
}

#[test]
fn test_result_map_chain_transitions() {
    let out = run_prints(r#"
        fun main() {
            val value = runCatching { 3 }
                .map { it * 2 }
                .map { it + 1 }
            println(value.getOrNull())
            println(value.isSuccess)
        }
    "#);
    assert_eq!(out, &["7", "true"]);
}

#[test]
fn test_result_chain_with_recover_and_map() {
    let out = run_prints(r#"
        fun main() {
            val value = runCatching { throw Exception("x") }
                .recover { 4 }
                .map { it + 1 }
            println(value.getOrNull())
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_result_nested_recover_chain() {
    let out = run_prints(r#"
        fun main() {
            val value = runCatching { 1 }
                .map { throw Exception("x") }
                .recover { 2 }
                .mapCatching { it + 1 }
            println(value.getOrNull())
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_result_exception_or_null_for_success_is_null() {
    let out = run_prints(r#"
        fun main() {
            val value = runCatching { 1 }
            println(value.exceptionOrNull() == null)
            println(value.getOrElse { -1 })
        }
    "#);
    assert_eq!(out, &["true", "1"]);
}

#[test]
fn test_result_success_with_null_value() {
    let out = run_prints(r#"
        fun main() {
            val value = Result.success<String?>(null)
            println(value.isSuccess)
            println(value.getOrNull() == null)
            println(value.getOrElse { "fallback" } ?: "none")
        }
    "#);
    assert_eq!(out, &["true", "true", "null"]);
}
