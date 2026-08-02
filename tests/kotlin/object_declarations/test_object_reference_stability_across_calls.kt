// vybe-test: kotlin/object_declarations/test_object_reference_stability_across_calls
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Counter {
            var value = 1
            fun reset() { value = 0 }
        }

        fun touch(): Int {
            Counter.value += 1
            return Counter.value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Counter.reset()
            __check((touch()).toString(), "1")
            __check((touch()).toString(), "2")
            __check((Counter.value).toString(), "2")
        }
