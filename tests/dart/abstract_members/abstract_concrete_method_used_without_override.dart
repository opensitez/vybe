// vybe-test: dart/abstract_members/abstract_concrete_method_used_without_override
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Logger {
  void log(String msg) {
    __p('log:$msg');
  }
  void flush();
}
class FileLogger extends Logger {
  void flush() {
    __p('flushed');
  }
}
void __vybeMain() {
  FileLogger().log('x');
}

void main() {
  __vybeMain();
  __check('log:x');
}
