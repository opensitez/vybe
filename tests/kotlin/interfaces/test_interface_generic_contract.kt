// vybe-test: kotlin/interfaces/test_interface_generic_contract
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Boxed<T> {
            val payload: T
            fun unwrap(): T
        }

        class IntBox(override val payload: Int) : Boxed<Int> {
            override fun unwrap(): Int = payload
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Boxed<Int> = IntBox(9)
            __check((value.unwrap()).toString(), "9")
            __check((value.payload).toString(), "9")
        }
