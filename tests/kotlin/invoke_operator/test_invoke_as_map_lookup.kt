// vybe-test: kotlin/invoke_operator/test_invoke_as_map_lookup
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Lookup {
            private val data = mapOf("a" to 1, "b" to 2)
            operator fun invoke(key: String): Int = data[key] ?: 0
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val l = Lookup()
            __check((l("a")).toString(), "1")
            __check((l("z")).toString(), "0")
        }
