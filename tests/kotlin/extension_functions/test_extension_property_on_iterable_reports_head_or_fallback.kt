// vybe-test: kotlin/extension_functions/test_extension_property_on_iterable_reports_head_or_fallback
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

val <T> Iterable<T>.headOrFallback: String
            get() = if (iterator().hasNext()) iterator().next().toString() else "fallback"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((listOf("a", "b").headOrFallback).toString(), "a")
            __check((listOf<Int>().headOrFallback).toString(), "fallback")
        }
