// vybe-test: kotlin/when_expressions/test_when_on_data_class_property_subject
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

data class User(val name: String, val active: Boolean, val level: Int)

        fun label(user: User): String {
            return when {
                user.name.isEmpty() -> "anon"
                !user.active -> "inactive"
                user.level > 10 -> "vip"
                else -> "regular"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(User("", true, 3))).toString(), "anon")
            __check((label(User("a", false, 1))).toString(), "inactive")
            __check((label(User("b", true, 12))).toString(), "vip")
            __check((label(User("c", true, 4))).toString(), "regular")
        }
