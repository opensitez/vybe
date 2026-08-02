// vybe-test: kotlin/object_declarations/test_singleton_object_holds_mutable_state
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Counter {
            var value = 0
            fun inc() { value += 1 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Counter.inc()
            Counter.inc()
            __check((Counter.value).toString(), "2")
        }
