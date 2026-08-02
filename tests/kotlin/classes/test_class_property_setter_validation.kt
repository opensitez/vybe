// vybe-test: kotlin/classes/test_class_property_setter_validation
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Counter {
            var value: Int = 0
                set(next) {
                    field = if (next < 0) 0 else next
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
            c.value = 5
            __check((c.value).toString(), "5")
            c.value = -3
            __check((c.value).toString(), "0")
        }
