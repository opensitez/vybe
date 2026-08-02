// vybe-test: kotlin/interfaces/test_interface_mixed_implementation
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface A {
            fun name(): String
        }

        interface B {
            fun count(): Int
        }

        class Combo(val label: String, val amount: Int) : A, B {
            override fun name(): String = label
            override fun count(): Int = amount
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Combo("item", 4)
            __check((c.name()).toString(), "item")
            __check((c.count()).toString(), "4")
        }
