// vybe-test: kotlin/collections/test_array_of_functions_dispatch
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ops = arrayOf(
                { x: Int -> x + 1 },
                { x: Int -> x * 2 },
                { x: Int -> x * x }
            )
            __check((ops[0](3)).toString(), "4")
            __check((ops[1](4)).toString(), "8")
            __check((ops[2](5)).toString(), "25")
        }
