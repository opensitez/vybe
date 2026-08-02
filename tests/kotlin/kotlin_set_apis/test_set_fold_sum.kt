// vybe-test: kotlin/kotlin_set_apis/test_set_fold_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf(1, 2, 3)
            __check((set.fold(0) { acc, v -> acc + v }).toString(), "6")
        }
