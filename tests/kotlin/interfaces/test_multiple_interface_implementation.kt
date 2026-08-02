// vybe-test: kotlin/interfaces/test_multiple_interface_implementation
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Named {
            fun getName(): String
        }

        interface Aged {
            fun getAge(): Int
        }

        class Citizen(val name: String, val age: Int) : Named, Aged {
            override fun getName(): String = name
            override fun getAge(): Int = age
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Citizen("Bob", 40)
            __check((c.getName()).toString(), "Bob")
            __check((c.getAge()).toString(), "40")
        }
