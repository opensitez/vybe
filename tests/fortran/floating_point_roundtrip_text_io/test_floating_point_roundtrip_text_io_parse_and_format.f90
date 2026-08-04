! vybe-test: fortran/floating_point_roundtrip_text_io/test_floating_point_roundtrip_text_io_parse_and_format
! origin: languages/fortran/tests/fortran/test_floating_point_roundtrip_text_io.rs

program test_floating_point_roundtrip_text_io
    real :: value
    character(len=20) :: text
    write(text, '(F8.3)') 1.5
    read(text, '(F8.3)') value
    print *, value
end program test_floating_point_roundtrip_text_io
