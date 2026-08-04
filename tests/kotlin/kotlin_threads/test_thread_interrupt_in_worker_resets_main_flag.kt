// vybe-test: kotlin/kotlin_threads/test_thread_interrupt_in_worker_resets_main_flag
// origin: languages/kotlin/tests/kotlin/test_kotlin_threads.rs

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
            val worker = kotlin.concurrent.thread(start = false) {
                try {
                    Thread.sleep(1000)
                    __p(("complete").toString())
                } catch (ex: InterruptedException) {
                    __p(("interrupted").toString())
                }
            }
            worker.start()
            worker.interrupt()
            worker.join()
            __p((worker.isInterrupted).toString())
            __p((worker.isInterrupted()).toString())
        
__check("interrupted\nfalse\nfalse")
}
