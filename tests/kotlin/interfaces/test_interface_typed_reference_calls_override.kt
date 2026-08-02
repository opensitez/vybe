// vybe-test: kotlin/interfaces/test_interface_typed_reference_calls_override
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Worker {
            fun work(): Int
        }

        class Engineer : Worker {
            override fun work(): Int = 3
        }

        fun report(w: Worker): Int {
            return w.work()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((report(Engineer())).toString(), "3")
        }
