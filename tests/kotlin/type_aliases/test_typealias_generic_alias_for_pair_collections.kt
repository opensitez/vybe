// vybe-test: kotlin/type_aliases/test_typealias_generic_alias_for_pair_collections
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias PairList<T> = List<Pair<T, T>>

        fun total(values: PairList<Int>): Int {
            return values.fold(0) { acc, item -> acc + item.first + item.second }
        }

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
            val values: PairList<Int> = listOf(Pair(1, 2), Pair(3, 4))
            __p((total(values)).toString())
        
__check("10")
}
