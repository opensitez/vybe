// vybe-test: kotlin/secondary_constructors/test_constructor_super_call
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

open class Animal(val name: String)

        class Dog : Animal {
            val age: Int

            constructor(name: String, age: Int) : super(name) {
                this.age = age
            }

            constructor(name: String) : this(name, 1)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Dog("Rex", 5)
            val b = Dog("Buddy")
            __check((a.name).toString(), "Rex")
            __check((a.age).toString(), "5")
            __check((b.name).toString(), "Buddy")
            __check((b.age).toString(), "1")
        }
