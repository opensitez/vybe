// vybe-test: kotlin/constructor_chaining/test_constructor_property_order
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Order {
            val a: Int
            val b: Int
            init { b = 2
a = 1 }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val o = Order()
            __check((o.a).toString(), "1")
            __check((o.b).toString(), "2")
        }
