// vybe-test: kotlin/properties/test_property_multiple_instances_are_isolated
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Tracker {
            var score: Int = 0
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Tracker()
            val b = Tracker()
            a.score = 4
            b.score = 9
            __check((a.score + b.score).toString(), "13")
        }
