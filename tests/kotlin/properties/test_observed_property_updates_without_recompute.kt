// vybe-test: kotlin/properties/test_observed_property_updates_without_recompute
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Item {
            private var count = 0
            val snapshot: Int
                get() = count

            var value: Int
                get() = count
                set(next) { count = next + 1 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Item()
            item.value = 7
            __check((item.snapshot).toString(), "8")
            item.value = 1
            __check((item.snapshot).toString(), "2")
            __check((item.value).toString(), "2")
        }
