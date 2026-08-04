! vybe-test: fortran/internal_io_extended/iio_read_formatted_a_substring
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf = 'alpha beta'
character(len=5) :: word
read(buf, '(A5)') word
if (trim(word) /= "alpha") then
    print *, "FAIL: want [alpha] got [", word, "]"
    stop 1
end if
end program t
