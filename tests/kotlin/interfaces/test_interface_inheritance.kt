// vybe-test: kotlin/interfaces/test_interface_inheritance
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Named {
            fun name(): String
        }

        interface Described : Named {
            fun description(): String
        }

        class Product : Described {
            override fun name(): String = "p"
            override fun description(): String = name() + "rod"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Product()
            __check((p.name()).toString(), "p")
            __check((p.description()).toString(), "prod")
        }
