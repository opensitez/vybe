! vybe-test: fortran/associate_construct_extended/associate_dtype_char_field
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
type :: Label
character(len=5) :: text
end type Label
type(Label) :: lbl
lbl%text = 'hello'
associate (msg => lbl%text)
if (trim(trim(msg)) /= "hello") then
    print *, "FAIL: want [hello] got [", trim(msg), "]"
    stop 1
end if
end associate
end program t
