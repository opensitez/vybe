// vybe-test: dart/comparable_ordering/comparable_priority_queue_style_insert
// origin: languages/dart/tests/dart/test_comparable_ordering.rs

final StringBuffer __vybeOut = StringBuffer();

void __p(Object? o) {
  __vybeOut.writeln(o);
}

void __check(String want) {
  var got = __vybeOut.toString();
  // `writeln` on the final print contributes a trailing newline that the
  // expected line vector never carried.
  if (got.endsWith('\n')) {
    got = got.substring(0, got.length - 1);
  }
  if (got != want) {
    print('FAIL: want [$want] got [$got]');
    throw Exception('assertion failed');
  }
}

class Task implements Comparable<Task> {
  int priority;
  Task(this.priority);
  int compareTo(Task other) => priority.compareTo(other.priority);
}
void __vybeMain() {
  var tasks = [Task(3), Task(1), Task(2)];
  tasks.sort();
  __p(tasks.first.priority);
}

void main() {
  __vybeMain();
  __check('1');
}
