! vybe-test: fortran/include_directive_line_continuations/test_include_directive_line_continuations_char_concat_split_across_lines
! origin: languages/fortran/tests/fortran/test_include_directive_line_continuations.rs

program test_include_directive_line_continuations
    character(len=20) :: word
    word = "for" // &
           "tran"
    if (trim(word) /= "fortran") then
    print *, "FAIL: want [fortran] got [", word, "]"
    stop 1
end if
end program test_include_directive_line_continuations
