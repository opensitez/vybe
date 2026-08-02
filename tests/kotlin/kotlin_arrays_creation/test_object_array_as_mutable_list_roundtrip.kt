// vybe-test: kotlin/kotlin_arrays_creation/test_object_array_as_mutable_list_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_arrays_creation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = arrayOf("a", "b", "c")
            val copy = values.toMutableList()
            copy.add("d")
            __check((copy.joinToString(";")).toString(), "a;b;c;d")
        }
