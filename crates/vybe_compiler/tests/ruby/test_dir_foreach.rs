use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dir_foreach_basic, "require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); File.write(\"#{d}/f.txt\", ''); puts Dir.foreach(d).to_a.sort.join('-') end", ".-..-f.txt-sub");
ruby_test!(test_dir_foreach_no_block, "require 'tmpdir'; Dir.mktmpdir {|d| puts Dir.foreach(d).is_a?(Enumerator)}", "true");
ruby_test!(test_dir_foreach_error, "begin; Dir.foreach('/non_existent_dir').to_a; rescue Errno::ENOENT; puts 'err'; end", "err");
ruby_test!(test_dir_entries_basic, "require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); puts Dir.entries(d).sort.join('-') end", ".-..-sub");
ruby_test!(test_dir_entries_error, "begin; Dir.entries('/non_existent_dir'); rescue Errno::ENOENT; puts 'err'; end", "err");
ruby_test!(test_dir_each_child_basic, "require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); puts Dir.each_child(d).to_a.sort.join('-') end", "sub"); // excludes . and ..
ruby_test!(test_dir_each_child_no_block, "require 'tmpdir'; Dir.mktmpdir {|d| puts Dir.each_child(d).is_a?(Enumerator)}", "true");
ruby_test!(test_dir_children_basic, "require 'tmpdir'; Dir.mktmpdir do |d| Dir.mkdir(\"#{d}/sub\"); puts Dir.children(d).sort.join('-') end", "sub");
