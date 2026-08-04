// vybe-test: dart/covariant_keyword/covariant_param_media_play
// origin: languages/dart/tests/dart/test_covariant_keyword.rs

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

class Media {}
class Audio extends Media {
  int duration;
  Audio(this.duration);
}
class Player {
  void play(Media m) {}
}
class AudioPlayer extends Player {
  @override
  void play(covariant Audio a) {
    __p(a.duration);
  }
}
void __vybeMain() {
  AudioPlayer().play(Audio(120));
}

void main() {
  __vybeMain();
  __check('120');
}
