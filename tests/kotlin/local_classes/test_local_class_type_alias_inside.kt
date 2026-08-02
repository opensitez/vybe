// vybe-test: kotlin/local_classes/test_local_class_type_alias_inside
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            typealias Text = String
            class Local(val v: Text)
            __check((Local("x").v).toString(), "x")
        }
