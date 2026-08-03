// Vybe test harness — Dart.
//
// Real Dart: this file analyses and runs under the Dart SDK on its own, which
// is what lets an extracted test be compared against the reference
// implementation.
//
// Output is COLLECTED, not paired. The emitter rewrites every `print(x)` in
// the test body into `__p(x)` and compares the whole output once, so a program
// whose print count is not static — a loop — still asserts. Pairing the i-th
// print with the i-th expected line cannot express that, and loops are the
// largest unpairable category in every language where it was tried.
//
// The test's own `main` is renamed to `__vybeMain` and called from a wrapper,
// which avoids having to find the closing brace of `main` to append the check.
// The wrapper is always `async` and always `await`s: `await` on a non-Future
// is legal Dart and returns the value unchanged, so one shape covers both
// `void main()` and `Future<void> main() async`.
//
// `print` here is the REAL print, deliberately: only the body's prints are
// rewritten, so the diagnostic still reaches stdout. It is emitted BEFORE the
// throw because an uncaught error renders as `RuntimeError: [object]` under
// Vybe, which would lose the expected and actual values entirely.

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
