// vybe-test: kotlin/interfaces/test_interface_implements_in_local_scope
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Callable {
            fun call(): Int
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Local : Callable {
                override fun call(): Int = 4
            }
            val item = Local()
            val c: Callable = item
            __check((c.call()).toString(), "4")
        }
