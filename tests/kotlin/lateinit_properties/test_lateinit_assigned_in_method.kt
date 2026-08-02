// vybe-test: kotlin/lateinit_properties/test_lateinit_assigned_in_method
// origin: languages/kotlin/tests/kotlin/test_lateinit_properties.rs

class Container {
            lateinit var value: String

            fun prepare() {
                value = "ready"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Container()
            c.prepare()
            __check((c.value).toString(), "ready")
        }
