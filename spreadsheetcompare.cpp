#include "spreadsheet.h"


//If ascending, what makes row1 < row2 should be true.
//else, what makes row1 > row2 should be true.
bool SpreadsheetCompare::operator()(const QStringList &row1,
	const QStringList &row2) const {
	for (int i = 0; i < KeyCount; ++i) {
		int column = keys[i];
		if (column != -1) {//If key != None.
			if (row1[column] != row2[column]) {
				if (ascending[i]) {
					return row1[column] < row2[column];
				}
				else {
					return row1[column] > row2[column];
				}
			}
		}
	}
	return false;
}

