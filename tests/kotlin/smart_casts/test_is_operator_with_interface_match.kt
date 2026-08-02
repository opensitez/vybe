// vybe-test: kotlin/smart_casts/test_is_operator_with_interface_match
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

interface Pet
        class Dog : Pet
        class Car

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pet: Any = Dog()
            val other: Any = Car()
            __check((pet is Pet).toString(), "true")
            __check((other is Pet).toString(), "false")
        }
