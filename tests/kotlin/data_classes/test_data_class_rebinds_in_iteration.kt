// vybe-test: kotlin/data_classes/test_data_class_rebinds_in_iteration
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Meter(val id: Int, val value: Int)

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
            val items = mutableListOf(Meter(1, 1), Meter(2, 2))
            var sum = 0
            for (item in items) {
                val updated = item.copy(value = item.value + 5)
                sum += updated.value
            }
            __p((sum).toString())
            __p((items[0].value).toString())
        
__check("13\n1")
}
