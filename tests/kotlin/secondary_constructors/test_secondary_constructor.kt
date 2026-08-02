// vybe-test: kotlin/secondary_constructors/test_secondary_constructor
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Person {
            val name: String
            constructor(name: String) {
                this.name = name
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Person("Alice")
            __check((p.name).toString(), "Alice")
        }
