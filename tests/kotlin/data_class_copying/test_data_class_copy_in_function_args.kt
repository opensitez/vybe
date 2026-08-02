// vybe-test: kotlin/data_class_copying/test_data_class_copy_in_function_args
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class User(val name: String, val level: Int)

        fun upgrade(user: User): User = user.copy(level = user.level + 1)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val user = User("x", 1)
            val next = upgrade(user)
            __check((next.name).toString(), "x")
            __check((next.level).toString(), "2")
        }
