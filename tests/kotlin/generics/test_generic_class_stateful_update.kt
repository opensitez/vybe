// vybe-test: kotlin/generics/test_generic_class_stateful_update
// origin: languages/kotlin/tests/kotlin/test_generics.rs

class Store<T>(start: T) {
            private var value: T = start

            fun get(): T = value
            fun set(next: T) {
                value = next
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = Store("a")
            val number = Store(10)
            text.set("b")
            number.set(11)
            __check((text.get()).toString(), "b")
            __check((number.get()).toString(), "11")
        }
