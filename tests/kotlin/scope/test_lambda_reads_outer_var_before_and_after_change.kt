// vybe-test: kotlin/scope/test_lambda_reads_outer_var_before_and_after_change
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 1
            val addOne = { total += 1 }
            addOne()
            __check((total).toString(), "2")
            total = 10
            addOne()
            __check((total).toString(), "11")
        }
