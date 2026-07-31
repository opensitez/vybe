use crate::helpers::run_prints;

#[test]
fn test_result_run_catching_success_path() {
    let out = run_prints(r#"
        fun main() {
            val result = runCatching { 5 + 7 }
            println(result.isSuccess)
            println(result.getOrNull())
            println(result.exceptionOrNull() == null)
        }
    "#);
    assert_eq!(out, &["true", "12", "true"]);
}

#[test]
fn test_result_run_catching_failure_path() {
    let out = run_prints(r#"
        fun main() {
            val result = runCatching { throw IllegalArgumentException("bad") }
            println(result.isFailure)
            println(result.isSuccess)
            println(result.exceptionOrNull()?.message)
            println(result.getOrNull() == null)
        }
    "#);
    assert_eq!(out, &["true", "false", "bad", "true"]);
}

#[test]
fn test_result_get_or_else_fallback() {
    let out = run_prints(r#"
        fun main() {
            val ok = runCatching { 10 }
            val fallback = ok.getOrElse { 0 }
            val bad = runCatching { throw IllegalStateException("x") }
            val fallbackBad = bad.getOrElse { 99 }
            println(fallback)
            println(fallbackBad)
        }
    "#);
    assert_eq!(out, &["10", "99"]);
}

#[test]
fn test_result_map_only_on_success() {
    let out = run_prints(r#"
        fun main() {
            val mapped = runCatching { 3 }
                .map { it * 10 }
                .map { it + 1 }
            val failed = runCatching<Int> { throw RuntimeException("fail") }
                .map { it + 1 }
            println(mapped.getOrNull())
            println(failed.isFailure)
            println(failed.getOrNull() == null)
        }
    "#);
    assert_eq!(out, &["31", "true", "true"]);
}

#[test]
fn test_result_map_catching_can_transform_failures() {
    let out = run_prints(r#"
        fun main() {
            val result = runCatching { "x".toInt() }
                .mapCatching { it + 1 }
            println(result.isFailure)
            val mapped = runCatching { 9 }
                .mapCatching { if (it % 2 == 1) throw IllegalArgumentException("odd") else it }
            println(mapped.isFailure)
            println(mapped.exceptionOrNull()?.message)
        }
    "#);
    assert_eq!(out, &["true", "true", "odd"]);
}

#[test]
fn test_result_recover_from_failure() {
    let out = run_prints(r#"
        fun main() {
            val result = runCatching { throw IllegalArgumentException("bad") }
                .recover { 12 }
            val failed = runCatching<Int> { throw IllegalStateException("boom") }
                .recover { cause -> if (cause is IllegalArgumentException) -1 else 99 }
            println(result.getOrNull())
            println(failed.getOrNull())
        }
    "#);
    assert_eq!(out, &["12", "99"]);
}

#[test]
fn test_result_recover_catching() {
    let out = run_prints(r#"
        fun main() {
            val result = runCatching { throw IllegalArgumentException("bad") }
                .recoverCatching { throw IllegalStateException("wrapped") }
            println(result.isFailure)
            println(result.exceptionOrNull()?.let { it::class.simpleName })
        }
    "#);
    assert_eq!(out, &["true", "IllegalStateException"]);
}

#[test]
fn test_result_fold_routes_success_and_failure() {
    let out = run_prints(r#"
        fun main() {
            val ok = runCatching { "7".toInt() }
                .fold(
                    onSuccess = { value -> "ok-" + value.toString() },
                    onFailure = { "bad" }
                )
            val bad = runCatching<Int> { "x".toInt() }
                .fold(
                    onSuccess = { value -> "ok-" + value.toString() },
                    onFailure = { e -> "err-" + e::class.simpleName.toString() }
                )
            println(ok)
            println(bad)
        }
    "#);
    assert_eq!(out, &["ok-7", "err-NumberFormatException"]);
}

#[test]
fn test_result_on_success_and_on_failure_have_side_effects() {
    let out = run_prints(r#"
        fun main() {
            var successSeen = false
            var failureSeen = false
            runCatching { 5 }
                .onSuccess { successSeen = true }
                .onFailure { failureSeen = true }
            runCatching<Int> { throw RuntimeException("fail") }
                .onSuccess { successSeen = true }
                .onFailure { failureSeen = true }
            println(successSeen)
            println(failureSeen)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_result_success_and_failure_factory() {
    let out = run_prints(r#"
        fun main() {
            val ok: Result<Int> = Result.success(21)
            val fail: Result<Int> = Result.failure(IllegalStateException("bad"))
            println(ok.getOrElse { -1 })
            println(fail.getOrElse { it.message?.length ?: 0 })
            println(ok.isSuccess)
            println(fail.isFailure)
        }
    "#);
    assert_eq!(out, &["21", "3", "true", "true"]);
}
