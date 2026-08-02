// vybe-test: kotlin/classes/test_class_declaration
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Person(val name: String, var age: Int) {
            fun greet() {
                __check(("I am " + name).toString(), "I am Alice")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Person("Alice", 30)
            p.greet()
        }
