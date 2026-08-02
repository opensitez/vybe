// vybe-test: kotlin/inner_classes/test_multiple_inner_instances_share_outer
// origin: languages/kotlin/tests/kotlin/test_inner_classes.rs

class Tracker {
            var ticks = 0
            inner class Probe {
                fun hit() { ticks += 1 }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Tracker()
            val a = t.Probe()
            val b = t.Probe()
            a.hit()
            b.hit()
            __check((t.ticks).toString(), "2")
        }
