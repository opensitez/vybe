// vybe-test: kotlin/property_accessors/test_property_setter_side_effect_log
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Logger {
            var value: Int = 0
                set(v) {
                    println(v)
                    field = v
                }
        }
        fun main() {
            val l = Logger()
            l.value = 1
            l.value = 2
        }

