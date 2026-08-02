// vybe-test: kotlin/data_classes/test_data_class_as_generic_type_argument
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Box<T>(val value: T)
        data class Holder<T>(val value: Box<T>)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val holder = Holder(Box("x"))
            __check((holder.value.value).toString(), "x")
            val copy = holder.copy(value = holder.value.copy(value = "y"))
            __check((holder.value.value).toString(), "x")
            __check((copy.value.value).toString(), "y")
        }
