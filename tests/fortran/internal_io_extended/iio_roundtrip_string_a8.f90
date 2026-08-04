! vybe-test: fortran/internal_io_extended/iio_roundtrip_string_a8
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf
character(len=4) :: s1 = 'vybe', s2
write(buf, '(A4)') s1
read(buf, '(A4)') s2
print *, s2
end program t
