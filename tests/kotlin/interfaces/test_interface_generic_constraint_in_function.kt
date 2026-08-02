// vybe-test: kotlin/interfaces/test_interface_generic_constraint_in_function
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Convertible<T> {
            fun convert(): T
        }

        class WrapInt : Convertible<String> {
            override fun convert(): String = "x"
        }

        fun <T> render(value: Convertible<T>): String {
            return value.convert().toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((render(WrapInt())).toString(), "x")
        }
