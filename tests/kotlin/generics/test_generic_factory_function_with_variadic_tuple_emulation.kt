// vybe-test: kotlin/generics/test_generic_factory_function_with_variadic_tuple_emulation
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <A, B> makePair(first: A, second: B): Array<Any?> {
            return arrayOf(first, second)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair1 = makePair(1, "k")
            val pair2 = makePair(true, 3.2)
            __check((pair1[0]).toString(), "1")
            __check((pair1[1]).toString(), "k")
            __check((pair2[0].toString() + pair2[1].toString()).toString(), "true3.2")
        }
