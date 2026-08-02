// vybe-test: kotlin/generics/test_generic_multiple_type_parameters
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <K, V> choose(left: K, right: V): String {
            return left.toString() + right.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((choose("a", 1)).toString(), "a1")
            __check((choose(2, "b")).toString(), "2b")
        }
