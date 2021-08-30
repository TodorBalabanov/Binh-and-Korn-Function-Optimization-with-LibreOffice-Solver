import uno
import random
from com.sun.star.sheet import SolverConstraint

def ParetoPointsGeneration():
	sheet = XSCRIPTCONTEXT.getDesktop().getCurrentComponent().CurrentController.ActiveSheet
	solver = uno.getComponentContext().ServiceManager.createInstanceWithContext("com.sun.star.sheet.Solver", uno.getComponentContext())
	solver.Document = XSCRIPTCONTEXT.getDocument()

	offset = int(sheet.getCellRangeByName("$M$2").getValue())

	lowerWeights = int(sheet.getCellRangeByName("$B$5").getValue())
	upperWeights = int(sheet.getCellRangeByName("$C$5").getValue())

	for g in range(lowerWeights,upperWeights+1):
		# Generate random weights of the objectives.
		sheet.getCellRangeByName("$G$2").setValue( g )
		for h in range(lowerWeights,upperWeights+1):
			# Generate random weights of the objectives.
			sheet.getCellRangeByName("$H$2").setValue( h )

			# Initialize random variables.
			sheet.getCellRangeByName("$B$2").setValue( random.uniform(float(sheet.getCellRangeByName("$B$3").getValue()),float(sheet.getCellRangeByName("$B$4").getValue())) )
			sheet.getCellRangeByName("$C$2").setValue( random.uniform(float(sheet.getCellRangeByName("$C$3").getValue()),float(sheet.getCellRangeByName("$C$4").getValue())) )

			solver.solve()

			# Store the values of the variables.
			sheet.getCellRangeByName("$B$" + str(offset)).setValue( sheet.getCellRangeByName("$B$2").getValue() )
			sheet.getCellRangeByName("$C$" + str(offset)).setValue( sheet.getCellRangeByName("$C$2").getValue() )

			# Store the constraints values.
			sheet.getCellRangeByName("$D$" + str(offset)).setValue( sheet.getCellRangeByName("$D$2").getValue() )
			sheet.getCellRangeByName("$E$" + str(offset)).setValue( sheet.getCellRangeByName("$E$2").getValue() )
			sheet.getCellRangeByName("$F$" + str(offset)).setValue( sheet.getCellRangeByName("$F$2").getValue() )

			# Store the weights used for the objectives.
			sheet.getCellRangeByName("$G$" + str(offset)).setValue( sheet.getCellRangeByName("$G$2").getValue() )
			sheet.getCellRangeByName("$H$" + str(offset)).setValue( sheet.getCellRangeByName("$H$2").getValue() )

			# Store the values of the objectives.
			sheet.getCellRangeByName("$I$" + str(offset)).setValue( sheet.getCellRangeByName("$I$2").getValue() )
			sheet.getCellRangeByName("$J$" + str(offset)).setValue( sheet.getCellRangeByName("$J$2").getValue() )

			offset = offset + 1

