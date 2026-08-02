// vybe-test: kotlin/kotlin_lazy_delegates/test_lazy_property_as_declaration_output
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_delegates.rs

class Holder {
            private val raw: Int = 3
            val doubled by lazy { raw * 2 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder()
            __check((h.doubled).toString(), "6")
            __check((h.doubled).toString(), "6")
        }
