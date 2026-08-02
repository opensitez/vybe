// vybe-test: kotlin/properties/test_property_companion_with_instance_and_class_read
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Counter {
            companion object {
                var next: Int = 0
            }

            fun take(): Int {
                Counter.next += 1
                return Counter.next
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c1 = Counter()
            val c2 = Counter()
            __check((c1.take()).toString(), "1")
            __check((c2.take()).toString(), "2")
        }
