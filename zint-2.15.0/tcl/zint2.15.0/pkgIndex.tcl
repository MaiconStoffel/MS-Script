if {[package vsatisfies [package provide Tcl] 9.0-]} { 
package ifneeded zint 2.15.0 [list load [file join $dir tcl9zint2150.dll]] 
} else { 
package ifneeded zint 2.15.0 [list load [file join $dir zint2150t.dll]] 
} 
