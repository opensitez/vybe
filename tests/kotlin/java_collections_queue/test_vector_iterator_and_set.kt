// vybe-test: kotlin/java_collections_queue/test_vector_iterator_and_set
// origin: languages/kotlin/tests/kotlin/test_java_collections_queue.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = java.util.Vector<Int>()
            v.add(1)
            v.add(2)
            v.addElement(3)
            val it = v.iterator()
            val sum = v.toMutableList().sum()
            __check((v.elementAt(1)).toString(), "2")
            __check((sum).toString(), "6")
            __check((it.next()).toString(), "1")
            __check((v.size).toString(), "3")
        }
