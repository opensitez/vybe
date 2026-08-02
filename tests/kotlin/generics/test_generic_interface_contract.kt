// vybe-test: kotlin/generics/test_generic_interface_contract
// origin: languages/kotlin/tests/kotlin/test_generics.rs

interface Provider<T> {
            fun get(): T
        }

        class Constant<T>(private val value: T) : Provider<T> {
            override fun get(): T = value
        }

        fun <T> read(provider: Provider<T>): T {
            return provider.get()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val name = Constant("Alice")
            val number = Constant(77)
            __check((read(name)).toString(), "Alice")
            __check((read(number)).toString(), "77")
        }
