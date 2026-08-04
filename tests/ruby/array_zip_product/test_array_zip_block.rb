# vybe-test: ruby/array_zip_product/test_array_zip_block
# origin: languages/ruby/tests/ruby/test_array_zip_product.rs

# Vybe test harness — Ruby.
#
# Real Ruby, exactly as test262's `assert.js` is real JavaScript: it runs under
# `/usr/bin/ruby` unchanged, which is what lets an extracted test be compared
# against the reference implementation.
#
# Unlike the Go/Kotlin/C# harnesses this does NOT pair the i-th print with the
# i-th expected line. It redefines `puts` to collect, so the test body is
# spliced in BYTE-IDENTICAL and its whole output is compared once. Two things
# fall out of that: loops work (the largest unpairable category everywhere
# else — 250 of 6,393 Go cases — is not a category here), and nothing rewrites
# the program, so nothing can corrupt it.
#
# Defining `puts` at top level defines a private method on Object, which is
# what shadows `Kernel#puts` — the same mechanism in real Ruby and in Vybe.
#
# NO `print` ANYWHERE IN THIS FILE, deliberately. Every other harness prints its
# own `FAIL: want [...] got [...]` before failing, because an uncaught error
# renders as `RuntimeError: [object]` and the message is otherwise lost. That is
# not possible here: Vybe's Ruby executes `print` inside a conditional
# unconditionally (measured — see project_ruby_print_breaks_conditionals), and
# the workaround `return if got == want` let a FAILING test exit 0. A harness
# that can silently pass is worse than one with no diagnostic, so the diagnostic
# goes and the verdict — the exit code, which is all testrunner reads — stays
# correct. Restore the FAIL line once `print` is fixed.

$__vybe_out = []

def puts(*args)
  if args.empty?
    # `puts` with no argument writes a single newline.
    $__vybe_out << ""
  else
    # `flatten` IS the semantics: real `puts` writes each element of an array
    # on its own line, recursively. It also avoids a type test — `is_a?` does
    # not work under Vybe, and `"str" + value` there renders `NaN`.
    args.flatten.each { |a| $__vybe_out << a.to_s }
  end
  nil
end

def __vybe_check(want)
  got = $__vybe_out.join("\n")
  if got != want
    raise "assertion failed"
  end
end

acc = []; [1, 2].zip([3, 4]) { |a, b| acc << a + b }; puts acc.join('-')

__vybe_check("4-6")
