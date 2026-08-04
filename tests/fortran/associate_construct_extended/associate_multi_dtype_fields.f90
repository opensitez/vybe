! vybe-test: fortran/associate_construct_extended/associate_multi_dtype_fields
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
type :: Rect
integer :: w, h
end type Rect
type(Rect) :: r
r%w = 4
r%h = 5
associate (width => r%w, height => r%h)
if ((width * height) /= 20) then
    print *, "FAIL: want [20] got [", width * height, "]"
    stop 1
end if
end associate
end program t
