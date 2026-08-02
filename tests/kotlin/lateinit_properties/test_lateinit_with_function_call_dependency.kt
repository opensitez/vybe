// vybe-test: kotlin/lateinit_properties/test_lateinit_with_function_call_dependency
// origin: languages/kotlin/tests/kotlin/test_lateinit_properties.rs

class Source {
            lateinit var source: String
        }

        fun build(s: Source): String = s.source

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = Source()
            s.source = "done"
            __check((build(s)).toString(), "done")
        }
