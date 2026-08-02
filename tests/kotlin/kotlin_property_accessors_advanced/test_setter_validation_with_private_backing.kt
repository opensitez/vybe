// vybe-test: kotlin/kotlin_property_accessors_advanced/test_setter_validation_with_private_backing
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class Counter {
            private var _count = 0
            var count: Int
                get() = _count
                set(v) {\n                    _count = if (v < 0) 0 else v
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
            c.count = -5
            __check((c.count).toString(), "0")
            c.count = 7
            __check((c.count).toString(), "7")
        }
