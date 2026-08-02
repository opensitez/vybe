// vybe-test: kotlin/data_classes/test_data_class_constructs_with_field_access
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class User(val name: String, val age: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = User("Ada", 30)
            __check((a.name).toString(), "Ada")
            __check((a.age).toString(), "30")
            __check((a.toString()).toString(), "User(name=Ada, age=30)")
        }
