kotlin_run_cases! {
    test_result_is_success_and_failure_flags => (r#"
        fun main() {
            val ok = runCatching { 10 }
            val bad = runCatching<Int> { throw IllegalArgumentException("bad") }
            println(ok.isSuccess)
            println(ok.isFailure)
            println(bad.isSuccess)
            println(bad.isFailure)
        }
    "#, &["true", "false", "false", "true"]),
    test_result_get_or_null_for_success => (r#"
        fun main() {
            val ok = runCatching { "7".toInt() }
            println(ok.getOrNull())
            println(ok.exceptionOrNull() == null)
        }
    "#, &["7", "true"]),
    test_result_get_or_null_for_failure => (r#"
        fun main() {
            val bad = runCatching<Int> { "x".toInt() }
            println(bad.getOrNull() == null)
            println(bad.exceptionOrNull()?.let { it::class.simpleName })
        }
    "#, &["true", "NumberFormatException"]),
    test_result_get_or_else_success => (r#"
        fun main() {
            val ok = runCatching { 5 }
            println(ok.getOrElse { 0 })
        }
    "#, &["5"]),
    test_result_get_or_else_failure => (r#"
        fun main() {
            val bad = runCatching<Int> { throw IllegalStateException("oops") }
            println(bad.getOrElse { 11 })
        }
    "#, &["11"]),
    test_result_get_or_default => (r#"
        fun main() {
            val ok = runCatching { 9 }
            val bad = runCatching<Int> { throw Exception("bad") }
            println(ok.getOrDefault(1))
            println(bad.getOrDefault(1))
        }
    "#, &["9", "1"]),
    test_result_recover_on_failure => (r#"
        fun main() {
            val bad = runCatching<String> { throw IllegalArgumentException("x") }
            val recovered = bad.recover { "fixed" }
            println(recovered.getOrNull())
            println(recovered.isFailure)
        }
    "#, &["fixed", "false"]),
    test_result_recover_success_passthrough => (r#"
        fun main() {
            val ok = runCatching { 3 }
            val recovered = ok.recover { -1 }
            println(recovered.getOrNull())
            println(recovered.isFailure)
        }
    "#, &["3", "false"]),
    test_result_recover_catching_flips_failure => (r#"
        fun main() {
            val bad = runCatching<Int> { throw IllegalArgumentException("bad") }
            val recovered = bad.recoverCatching { throw IllegalStateException("wrapped") }
            println(recovered.isFailure)
            println(recovered.exceptionOrNull()?.let { it::class.simpleName })
        }
    "#, &["true", "IllegalStateException"]),
    test_result_map_success_only => (r#"
        fun main() {
            val ok = runCatching { 4 }
                .map { it * 10 }
            val bad = runCatching<Int> { throw Exception("x") }
                .map { it * 2 }
            println(ok.getOrNull())
            println(bad.getOrNull() == null)
        }
    "#, &["40", "true"]),
    test_result_map_catching_success => (r#"
        fun main() {
            val ok = runCatching { 5 }
                .mapCatching { if (it == 5) "five" else throw Exception("x") }
            println(ok.getOrNull())
            println(ok.exceptionOrNull() == null)
        }
    "#, &["five", "true"]),
    test_result_map_catching_failure => (r#"
        fun main() {
            val bad = runCatching { "x".toInt() }
                .mapCatching { throw IllegalStateException("mapped") }
            println(bad.getOrNull() == null)
            println(bad.exceptionOrNull()?.message)
        }
    "#, &["true", "mapped"]),
    test_result_fold_success_and_failure => (r#"
        fun main() {
            val ok = runCatching { 2 }
                .fold({ it.toString() }, { "bad" })
            val bad = runCatching<Int> { throw RuntimeException("boom") }
                .fold({ it.toString() }, { e -> e::class.simpleName.toString() })
            println(ok)
            println(bad)
        }
    "#, &["2", "RuntimeException"]),
    test_result_on_success_side_effect => (r#"
        fun main() {
            var marker = ""
            runCatching { 4 }
                .onSuccess { marker += "ok" + it.toString() }
                .onFailure { marker += "fail" }
            println(marker)
        }
    "#, &["ok4"]),
    test_result_on_failure_side_effect => (r#"
        fun main() {
            var marker = ""
            runCatching<Int> { throw IllegalArgumentException("x") }
                .onSuccess { marker += "ok" }
                .onFailure { marker += it::class.simpleName.toString() }
            println(marker)
        }
    "#, &["IllegalArgumentException"]),
    test_result_on_success_then_failure_state => (r#"
        fun main() {
            var count = 0
            runCatching { 7 }.onSuccess { count += it }
            runCatching<Int> { throw RuntimeException("bad") }.onFailure { count += 5 }
            println(count)
        }
    "#, &["12"]),
    test_result_or_null_default_int => (r#"
        fun main() {
            val ok = runCatching { 8 }
            val bad = runCatching<Int> { throw Exception("n") }
            println(ok.getOrNull())
            println(bad.getOrNull() ?: 100)
        }
    "#, &["8", "100"]),
    test_result_fold_return_type_change => (r#"
        fun main() {
            val a = runCatching { 9 }
            val out = a.fold(
                onSuccess = { v -> v + 1 },
                onFailure = { _ -> 0 }
            )
            println(out)
        }
    "#, &["10"]),
    test_result_exception_message_chain => (r#"
        fun main() {
            val value = runCatching<Int> { 1 / 0 }
                .map { it + 1 }
                .recover { 0 }
            val thrown = runCatching<Int> { 1 / 0 }
                .exceptionOrNull()
            println(value)
            println(thrown?.let { it.message })
        }
    "#, &["0", "/ by zero"]),
    test_result_nested_run_catching => (r#"
        fun main() {
            val value = runCatching {
                runCatching { "3".toInt() }.getOrThrow()
            }
            println(value.isSuccess)
            println(value.getOrNull())
        }
    "#, &["true", "3"]),
    test_result_nested_failure_bubble => (r#"
        fun main() {
            val value = runCatching {
                runCatching<Int> { "x".toInt() }.getOrThrow()
            }
            println(value.isFailure)
            println(value.exceptionOrNull()?.let { it::class.simpleName })
        }
    "#, &["true", "NumberFormatException"]),
    test_result_recover_from_nested_failure => (r#"
        fun main() {
            val value = runCatching {
                runCatching<Int> { "x".toInt() }
            }.recover { Result.failure<Int>(RuntimeException("bad")) }
            println(value.isSuccess)
        }
    "#, &["true"]),
    test_result_result_success_factory => (r#"
        fun main() {
            val value: Result<String> = Result.success("ok")
            println(value.isSuccess)
            println(value.getOrNull())
        }
    "#, &["true", "ok"]),
    test_result_result_failure_factory => (r#"
        fun main() {
            val value: Result<String> = Result.failure(IllegalStateException("bad"))
            println(value.isFailure)
            println(value.exceptionOrNull()?.message)
        }
    "#, &["true", "bad"]),
    test_result_result_nullable_unwrap => (r#"
        fun main() {
            val value = runCatching { null as String? }
            val payload = value.getOrNull()
            println(payload == null)
        }
    "#, &["true"]),
    test_result_mixed_chain_map_then_recover => (r#"
        fun main() {
            val value = runCatching<Int> { "x".toInt() }
                .map { it * 2 }
                .recover { 21 }
            println(value.getOrElse { 0 })
        }
    "#, &["21"]),
    test_result_map_catching_then_recover => (r#"
        fun main() {
            val value = runCatching { 1 }
                .mapCatching { throw IllegalArgumentException("nope") }
                .recover { 99 }
            println(value.getOrNull())
        }
    "#, &["99"]),
    test_result_fold_failure_message => (r#"
        fun main() {
            val value = runCatching<Int> { 1 / 0 }
                .fold({ "ok" }, { e -> e::class.simpleName.toString() })
            println(value)
        }
    "#, &["ArithmeticException"]),
    test_result_recover_with_type_match => (r#"
        fun main() {
            val value = runCatching<Int> { throw IllegalArgumentException("x") }
                .recover { cause -> if (cause is IllegalArgumentException) 10 else 0 }
            println(value.getOrNull())
        }
    "#, &["10"]),
    test_result_recover_with_unrelated_type => (r#"
        fun main() {
            val value = runCatching<Int> { throw IllegalArgumentException("x") }
                .recover { cause -> if (cause is UnsupportedOperationException) 1 else 2 }
            println(value.getOrElse { 0 })
        }
    "#, &["2"]),
    test_result_is_failure_then_map_not_run => (r#"
        fun main() {
            var seen = 0
            val value = runCatching<Int> { throw Exception("x") }
                .map {
                    seen = 1
                    it
                }
            println(seen)
            println(value.isFailure)
        }
    "#, &["0", "true"]),
}
