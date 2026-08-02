// vybe-test: kotlin/extension_properties/test_extension_property_person_age_group
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

class Person(val age: Int)
        val Person.group: String get() = if (age < 18) "minor" else "adult"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Person(12).group).toString(), "minor")
            __check((Person(28).group).toString(), "adult")
        }
