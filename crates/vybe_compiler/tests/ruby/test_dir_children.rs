use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dir_children_basic, "require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); File.write(\"#{d}/f.txt\", ''); puts Dir.children(d).sort.join('-') end", "f.txt-sub");
ruby_test!(test_dir_children_empty, "require 'tmpdir'; Dir.mktmpdir do |d| puts Dir.children(d).length end", "0");
ruby_test!(test_dir_children_error, "begin; Dir.children('/non_existent_dir'); rescue Errno::ENOENT; puts 'err'; end", "err");
ruby_test!(test_dir_each_child_basic, "require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); File.write(\"#{d}/f.txt\", ''); puts Dir.each_child(d).to_a.sort.join('-') end", "f.txt-sub");
ruby_test!(test_dir_each_child_error, "begin; Dir.each_child('/non_existent_dir').to_a; rescue Errno::ENOENT; puts 'err'; end", "err");
