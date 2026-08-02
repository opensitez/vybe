// vybe-test: kotlin/kotlin_sorting_comparators/test_fold_right_concatenate
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", "b", "c").sorted()
            val out = values.foldRight("") { value, acc -> value + acc }
            __check((out).toString(), "abc")
        }
