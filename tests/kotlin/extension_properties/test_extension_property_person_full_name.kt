// vybe-test: kotlin/extension_properties/test_extension_property_person_full_name
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

class Person(val first: String, val last: String)
        val Person.fullName: String get() = "$first $last"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Person("Ada", "Lovelace").fullName).toString(), "Ada Lovelace")
        }
