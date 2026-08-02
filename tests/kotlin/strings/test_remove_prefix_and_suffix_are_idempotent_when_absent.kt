// vybe-test: kotlin/strings/test_remove_prefix_and_suffix_are_idempotent_when_absent
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val word = "kotlin"
            __check((word.removePrefix("ko")).toString(), "tlin")
            __check((word.removePrefix("x")).toString(), "kotlin")
            __check((word.removeSuffix("in")).toString(), "kotl")
            __check((word.removeSuffix("x")).toString(), "kotlin")
        }
