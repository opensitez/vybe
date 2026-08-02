// vybe-test: kotlin/kotlin_accessor_customization/test_property_init_before_setter
// origin: languages/kotlin/tests/kotlin/test_kotlin_accessor_customization.rs

class Counter {
            var value = 1
                private set

            fun setPublic(v: Int) {
                value = v
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter()
            c.setPublic(5)
            __check((c.value).toString(), "5")
        }
