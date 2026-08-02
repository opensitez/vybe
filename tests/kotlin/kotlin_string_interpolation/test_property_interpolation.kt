// vybe-test: kotlin/kotlin_string_interpolation/test_property_interpolation
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation.rs

class User(val name: String, val age: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = User("bob", 5)
            __check(("${u.name}:${u.age}").toString(), "bob:5")
        }
