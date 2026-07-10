use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_readlines_basic, "require 'tempfile'; t = Tempfile.new('rl'); t.write(\"a\\nb\\nc\"); t.rewind; puts t.readlines.map(&:strip).join('-')", "a-b-c");
ruby_test!(test_readlines_chomp, "require 'tempfile'; t = Tempfile.new('rl'); t.write(\"a\\nb\\nc\"); t.rewind; puts t.readlines(chomp: true).join('-')", "a-b-c");
ruby_test!(test_readlines_separator, "require 'tempfile'; t = Tempfile.new('rl'); t.write(\"a,b,c\"); t.rewind; puts t.readlines(',').join('-')", "a,-b,-c");
ruby_test!(test_readlines_limit, "require 'tempfile'; t = Tempfile.new('rl'); t.write(\"hello\\nworld\"); t.rewind; puts t.readlines(3).map(&:strip).join('-')", "hel-lo-wor-ld"); // limit characters per line
ruby_test!(test_readlines_separator_limit, "require 'tempfile'; t = Tempfile.new('rl'); t.write(\"hello,world\"); t.rewind; puts t.readlines(',', 3).map(&:strip).join('-')", "hel-lo,-wor-ld");
