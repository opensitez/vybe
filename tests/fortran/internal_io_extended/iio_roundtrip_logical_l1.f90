! vybe-test: fortran/internal_io_extended/iio_roundtrip_logical_l1
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=4) :: buf
logical :: a = .true., b
write(buf, '(L1)') a
read(buf, '(L1)') b
print *, b
end program t
