// vybe-test: kotlin/kotlin_sorting_comparators/test_reduce_with_indexed_variant
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(4, 5, 6).sorted()
            val out = values.foldIndexed("") { index, acc, value ->
                if (index == 0) value.toString() else "${'$'}{acc}-${'$'}value"
            }
            __check((out).toString(), "4-5-6")
        }
