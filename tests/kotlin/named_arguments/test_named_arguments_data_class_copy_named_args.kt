// vybe-test: kotlin/named_arguments/test_named_arguments_data_class_copy_named_args
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

data class User(val id: Int, val role: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = User(id = 2, role = "x")
            val v = u.copy(role = "admin")
            __check((v.id).toString(), "2")
            __check((v.role).toString(), "admin")
        }
