// vybe-test: kotlin/properties/test_property_primary_constructor_value_access
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class User(val name: String, val age: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val user = User("Ari", 27)
            __check((user.name).toString(), "Ari")
            __check((user.age).toString(), "27")
        }
