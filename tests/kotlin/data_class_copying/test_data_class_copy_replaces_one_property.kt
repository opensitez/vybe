// vybe-test: kotlin/data_class_copying/test_data_class_copy_replaces_one_property
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class User(val name: String, val age: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = User("a", 1)
            val b = a.copy(age = 2)
            __check((a.name).toString(), "a")
            __check((b.name).toString(), "a")
            __check((b.age).toString(), "2")
        }
