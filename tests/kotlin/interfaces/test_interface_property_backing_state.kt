// vybe-test: kotlin/interfaces/test_interface_property_backing_state
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Counter {
            var count: Int
        }

        class Stateful : Counter {
            private var backing = 1
            override var count: Int
                get() = backing
                set(value) { backing = value }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c: Counter = Stateful()
            __check((c.count).toString(), "1")
            c.count += 4
            __check((c.count).toString(), "5")
        }
