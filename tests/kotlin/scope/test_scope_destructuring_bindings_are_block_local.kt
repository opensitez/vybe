// vybe-test: kotlin/scope/test_scope_destructuring_bindings_are_block_local
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (first, second) = Pair("left", "right")

            val inner = run {
                val (first, second) = Pair("inner-left", "inner-right")
                first + ":" + second
            }

            __check((first).toString(), "left")
            __check((second).toString(), "right")
            __check((inner).toString(), "inner-left:inner-right")
        }
