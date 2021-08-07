import uno
import random
from com.sun.star.sheet import SolverConstraint
#from com.sun.star.sheet import SolverConstraintOperator

def ParetoPointsGeneration():
	LE = 0
	GE = 2

	document = XSCRIPTCONTEXT.getDocument()
	sheet = XSCRIPTCONTEXT.getDesktop().getCurrentComponent().CurrentController.ActiveSheet
	solver = uno.getComponentContext().ServiceManager.createInstanceWithContext("com.sun.star.sheet.Solver", uno.getComponentContext())

	solver.Document = document
#	solver.Objective = sheet.getCellRangeByName("$J$3").getCellAddress()
#	solver.Variables = [sheet.getCellRangeByName("$B$3").getCellAddress(), sheet.getCellRangeByName("$C$3").getCellAddress()]
#	solver.Constraints = [ 
#		SolverConstraint(sheet.getCellRangeByName("$B$3").getCellAddress(), GE, sheet.getCellRangeByName("$B$4").getCellAddress()),
#		SolverConstraint(sheet.getCellRangeByName("$B$3").getCellAddress(), LE, sheet.getCellRangeByName("$B$5").getCellAddress()),
#		SolverConstraint(sheet.getCellRangeByName("$C$3").getCellAddress(), GE, sheet.getCellRangeByName("$C$4").getCellAddress()),
#		SolverConstraint(sheet.getCellRangeByName("$C$3").getCellAddress(), LE, sheet.getCellRangeByName("$C$5").getCellAddress()),
#		SolverConstraint(sheet.getCellRangeByName("$E$3").getCellAddress(), GE, sheet.getCellRangeByName("$E$4").getCellAddress()),
#		SolverConstraint(sheet.getCellRangeByName("$D$3").getCellAddress(), LE, sheet.getCellRangeByName("$D$5").getCellAddress()),
#	]
#	solver.Maximize = False

	loops = int(sheet.getCellRangeByName("$K$3").getValue())
	offset = int(sheet.getCellRangeByName("$L$3").getValue())
	for g in range(0, loops):
		# Generate random weights of the objectives.
		sheet.getCellRangeByName("$F$4").setValue( random.randint(0,1000) )
		sheet.getCellRangeByName("$G$4").setValue( random.randint(0,1000) )
		sheet.getCellRangeByName("$B$3").setValue( random.random() )
		sheet.getCellRangeByName("$C$3").setValue( random.random() )

		solver.solve()

		# Store the values of the variables.
		sheet.getCellRangeByName("B" + str(offset)).setValue( sheet.getCellRangeByName("$B$3").getValue() )
		sheet.getCellRangeByName("C" + str(offset)).setValue( sheet.getCellRangeByName("$C$3").getValue() )

		# Store the weights used for the objectives.
		sheet.getCellRangeByName("F" + str(offset)).setValue( sheet.getCellRangeByName("$F$3").getValue() )
		sheet.getCellRangeByName("G" + str(offset)).setValue( sheet.getCellRangeByName("$G$3").getValue() )

		# Store the values of the objectives.
		sheet.getCellRangeByName("H" + str(offset)).setValue( sheet.getCellRangeByName("$H$3").getValue() )
		sheet.getCellRangeByName("I" + str(offset)).setValue( sheet.getCellRangeByName("$I$3").getValue() )

		offset = offset + 1
