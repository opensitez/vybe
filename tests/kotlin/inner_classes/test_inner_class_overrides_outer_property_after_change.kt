// vybe-test: kotlin/inner_classes/test_inner_class_overrides_outer_property_after_change
// origin: languages/kotlin/tests/kotlin/test_inner_classes.rs

class Counter {
            var base = 1
            inner class Ticker {
                fun tick() { base += 2 }
            }

            fun value(): Int = base
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter()
            val t = c.Ticker()
            t.tick()
            t.tick()
            __check((c.value()).toString(), "5")
        }
