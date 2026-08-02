// vybe-test: kotlin/properties/test_property_override_mutable_readwrite_property
// origin: languages/kotlin/tests/kotlin/test_properties.rs

interface CounterLike {
            var count: Int
        }

        class Stateful : CounterLike {
            private var raw = 1
            override var count: Int
                get() = raw
                set(next) { raw = next + 1 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c: CounterLike = Stateful()
            c.count = 2
            __check((c.count).toString(), "3")
        }
