// vybe-test: kotlin/string_templates/test_template_property_accessor
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

class User(val name: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = User("a")
            __check(("user=${u.name}").toString(), "user=a")
        }
