// vybe-test: kotlin/kotlin_constructor_parameters/test_defaulted_constructor_param
// origin: languages/kotlin/tests/kotlin/test_kotlin_constructor_parameters.rs

class Person(val name: String, val age: Int = 10) {
            fun describe(): String {
                return name + ":" + age.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Person("a")
            val b = Person("b", 20)
            __check((a.describe()).toString(), "a:10")
            __check((b.describe()).toString(), "b:20")
        }
