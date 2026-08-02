// vybe-test: kotlin/kotlin_accessor_customization/test_backed_property_with_side_effect
// origin: languages/kotlin/tests/kotlin/test_kotlin_accessor_customization.rs

class Store {
            var marker: String = ""
                set(v) {
                    field = v + "!"
                }
                get() = field + "#"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = Store()
            s.marker = "go"
            __check((s.marker).toString(), "go!#")
        }
