use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_ungetc_basic, "require 'tempfile'; t = Tempfile.new('ug'); t.write('hello'); t.rewind; t.getc; t.ungetc('A'); puts t.read", "Aello");
ruby_test!(test_ungetc_string, "require 'tempfile'; t = Tempfile.new('ug'); t.write('hello'); t.rewind; t.getc; t.ungetc('AB'); puts t.read", "ABello");
ruby_test!(test_ungetc_eof, "require 'tempfile'; t = Tempfile.new('ug'); t.ungetc('A'); puts t.read", "A"); // ungetc works even at EOF
ruby_test!(test_ungetc_pos, "require 'tempfile'; t = Tempfile.new('ug'); t.write('hello'); t.rewind; t.getc; t.ungetc('A'); puts t.pos", "0"); // ungetc pushes pos back
ruby_test!(test_ungetbyte_basic, "require 'tempfile'; t = Tempfile.new('ug'); t.write('hello'); t.rewind; t.getc; t.ungetbyte(65); puts t.read", "Aello"); // 65 is 'A'
ruby_test!(test_ungetbyte_string, "require 'tempfile'; t = Tempfile.new('ug'); t.write('hello'); t.rewind; t.getc; t.ungetbyte('AB'); puts t.read", "ABello");
ruby_test!(test_ungetbyte_eof, "require 'tempfile'; t = Tempfile.new('ug'); t.ungetbyte(65); puts t.read", "A");
