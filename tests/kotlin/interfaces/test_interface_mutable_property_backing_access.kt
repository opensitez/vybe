// vybe-test: kotlin/interfaces/test_interface_mutable_property_backing_access
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Counter {
            var value: Int
        }

        class Store(initial: Int) : Counter {
            override var value: Int = initial
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c: Counter = Store(2)
            __check((c.value).toString(), "2")
            c.value += 5
            __check((c.value).toString(), "7")
        }
