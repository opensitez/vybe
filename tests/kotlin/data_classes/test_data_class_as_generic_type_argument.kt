// vybe-test: kotlin/data_classes/test_data_class_as_generic_type_argument
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Box<T>(val value: T)
        data class Holder<T>(val value: Box<T>)

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val holder = Holder(Box("x"))
            __p((holder.value.value).toString())
            val copy = holder.copy(value = holder.value.copy(value = "y"))
            __p((holder.value.value).toString())
            __p((copy.value.value).toString())
        
__check("x\nx\ny")
}
