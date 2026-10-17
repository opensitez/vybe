fun kotlin_controls(): Int {
    val s = mutableSetOf(1, 2, 3)
    val a = s.add(4)
    val b = s.add(4)
    val c = s.remove(2)
    val d = s.remove(9)
    return s.size * 16 + (if (a) 8 else 0) + (if (b) 4 else 0) + (if (c) 2 else 0) + (if (d) 1 else 0)
}

fun kotlin_cycle(rounds: Int): Int {
    val s = mutableSetOf(1, 2, 3)
    var i = 0
    while (i < rounds) {
        s.remove(2)
        s.add(2)
        i++
    }
    return s.size
}
