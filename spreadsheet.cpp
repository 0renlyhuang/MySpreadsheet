#include <qabstractitemview.h>
#include <qmessagebox.h>
#include <qfile.h>
#include <qapplication.h>
#include <qclipboard.h>

#include "spreadsheet.h"
#include "cell.h"

Spreadsheet::Spreadsheet(QWidget *parent)
	: QTableWidget(parent) {
	autoRecalc = true;//Should wirte in spreadsheet.h.

	//The table widget will use the cell's clone function 
	//when it needs to create a new table item
	setItemPrototype(new Cell);
	//The cells can be selected by dragging a range with the mouse
	setSelectionMode(ContiguousSelection);


	connect(this, SIGNAL(itemChanged(QTableWidgetItem*)),
		this, SLOT(somethingChanged()));

	clear();
}

//Connected with StatusBar.
QString Spreadsheet::currentLocation() const {
	return QChar('A' + currentColumn())
		+ QString::number(currentRow() + 1);
}

//Connected with StatusBar.
QString Spreadsheet::currentFormula() const {
	return formula(currentRow(), currentColumn());
}

QTableWidgetSelectionRange Spreadsheet::selectedRange() const {
	//Using function selectedRanges() to return the selected ranges 
	//no matter in what ways they were selected.	

	//In fact, ranges will never be empty, 
	//because the selectionMode ContiguousSelection will assume the current cell 
	//as the selected cell. 	
	QList<QTableWidgetSelectionRange> ranges = selectedRanges();
	if (ranges.isEmpty())
		return QTableWidgetSelectionRange();
	return ranges.first();
}


void Spreadsheet::clear() {
	setRowCount(0);
	setColumnCount(0);//Clear the whole spreadsheet.
	setRowCount(RowCount);
	setColumnCount(ColumnCount);

	for (int i = 0; i < ColumnCount; ++i) {
		QTableWidgetItem *item = new QTableWidgetItem;
		item->setText(QString(QChar('A' + i)));
		setHorizontalHeaderItem(i, item);
	}

	setCurrentCell(0, 0);
}

bool Spreadsheet::readFile(const QString &fileName) {
	QFile file(fileName);
	if (!file.open(QIODevice::ReadOnly)) {
		QMessageBox::warning(this, tr("Spreadsheet"),
			tr("Cannot read file %1:\n%2.")
			.arg(file.fileName())
			.arg(file.errorString()));
		return false;
	}
	QDataStream in(&file);
	in.setVersion(QDataStream::Qt_5_5);

	quint32 magic;
	in >> magic;
	if (magic != MagicNumber) {
		QMessageBox::warning(this, tr("Spreadsheet"),
			tr("This file isn't a spreadsheet file."));
		return false;
	}
	clear();

	quint16 row;
	quint16 column;
	QString str;

	QApplication::setOverrideCursor(Qt::WaitCursor);
	while (!in.atEnd()) {
		in >> row >> column >> str;
		setFormula(row, column, str);
	}
	QApplication::restoreOverrideCursor();
	return true;
}


bool Spreadsheet::writeFile(const QString &fileName) {
	QFile file(fileName);
	if (!file.open(QIODevice::WriteOnly)) {
		QMessageBox::warning(this, tr("Spreadsheet"),
			tr("Cannot write file %1\n%2.")
			.arg(file.fileName())
			.arg(file.errorString()));
		return false;
	}
	QDataStream out(&file);
	out.setVersion(QDataStream::Qt_5_5);

	out << quint32(MagicNumber);

	QApplication::setOverrideCursor(Qt::WaitCursor);
	QString str;
	for (int row = 0; row < RowCount; ++row) {
		for (int column = 0; column < ColumnCount; ++column) {
			str = formula(row, column);
			if (!str.isEmpty())
				out << quint16(row) << quint16(column) << str;
		}
	}
	QApplication::restoreOverrideCursor();
	return true;
}

void Spreadsheet::sort(const SpreadsheetCompare &compare) {
	QList<QStringList> rows;
	QTableWidgetSelectionRange range = selectedRange();

	for (int i = 0; i < range.rowCount(); ++i) {
		QStringList row;
		for (int j = 0; j < range.columnCount(); ++j) {
			row.append(formula(range.topRow() + i, range.leftColumn() + j));
		}
		rows.append(row);
	}

	qStableSort(rows.begin(), rows.end(), compare);

	for (int i = 0; i < range.rowCount(); ++i) {
		for (int j = 0; j < range.columnCount(); ++j) {
			setFormula(range.topRow() + i, range.leftColumn() + j, rows[i][j]);
		}
	}

	clearSelection();
	somethingChanged();
}
void Spreadsheet::cut() {
	copy();
	del();
}

void Spreadsheet::copy() {
	QTableWidgetSelectionRange range = selectedRange();
	QString str;

	for (int i = 0; i < range.rowCount(); ++i) {
		if (i > 0)
			str += "\n";
		for (int j = 0; j < range.columnCount(); ++j) {
			if (j > 0)
				str += "\t";
			str += formula(range.topRow() + i, range.leftColumn() + j);
		}
	}
	QApplication::clipboard()->setText(str);
}

void Spreadsheet::paste() {
	QTableWidgetSelectionRange range = selectedRange();
	QString str = QApplication::clipboard()->text();
	QStringList rows = str.split('\n');
	int numRows = rows.count();
	int numColunms = rows.first().count('\t') + 1;

	if (range.rowCount() * range.columnCount() != 1 //Effect remains unknowned.
		&& (range.rowCount() != numRows
		|| range.columnCount() != numColunms)) {
		QMessageBox::information(this, tr("MySpreadsheet"),
			tr("The information cannot be pasted because the copy"
			"and paste areas aren't the same size."));
		return;
	}

	QStringList columns;
	int row, column;
	for (int i = 0; i != numRows; ++i) {
		columns = rows[i].split('\t');
		for (int j = 0; j != numColunms; ++j) {
			row = range.topRow() + i;
			column = range.leftColumn() + j;
			if (row < RowCount && column < ColumnCount)
				setFormula(row, column, columns[j]);
		}
	}
	somethingChanged();
}

void Spreadsheet::del() {
	QList<QTableWidgetItem *> items = selectedItems();
	if (!items.isEmpty()) {
		foreach(QTableWidgetItem *item, items)
			delete item;
		somethingChanged();
	}
}

void Spreadsheet::selectCurrentRow() {
	selectRow(currentRow());
}

void Spreadsheet::selectCurrentColumn() {
	selectColumn(currentColumn());
}

void Spreadsheet::recalculate() {
	for (int row = 0; row < RowCount; ++row) {
		for (int column = 0; column < ColumnCount; ++column) {
			if (cell(row, column))
				cell(row, column)->setDirty();
		}
	}
	viewport()->update();
}

void Spreadsheet::setAutoRecalculate(bool recalc) {
	autoRecalc = recalc;
	if (autoRecalc)
		recalculate();
}

void Spreadsheet::findNext(const QString &str, Qt::CaseSensitivity cs) {
	int row = currentRow();
	int column = currentColumn() + 1;

	while (row < RowCount) {
		while (column < ColumnCount) {
			if (text(row, column).contains(str, cs)) {
				clearSelection();
				setCurrentCell(row, column);
				activateWindow();//Activates spreadsheet window(from finddialog).
				return;
			}
			++column;
		}
		column = 0;
		++row;
	}
	QApplication::beep();//When the software can not find the content.
}

void Spreadsheet::findPrevious(const QString &str, Qt::CaseSensitivity cs) {
	int row = currentRow();
	int column = currentColumn() - 1;

	while (row >= 0) {
		while (column >= 0) {
			if (text(row, column).contains(str, cs)) {
				clearSelection();
				setCurrentCell(row, column);
				activateWindow();
			}
			--column;
		}
		column = ColumnCount - 1;
		--row;
	}
	QApplication::beep();
}

void Spreadsheet::somethingChanged() {
	if (autoRecalc)
		recalculate();
	emit modified();
}


Cell *Spreadsheet::cell(int row, int column) const {
	return static_cast<Cell*>(item(row, column));
}


void Spreadsheet::setFormula(int row, int column, const QString &formula) {
	Cell *c = cell(row, column);
	if (!c) {
		c = new Cell;
		setItem(row, column, c);//Sets c to the table's item which locates at (row, column).
	}
	c->setFormula(formula);
}

QString Spreadsheet::formula(int row, int column) const {
	Cell *c = cell(row, column);
	if (c) {
		return c->formula();
	}
	else {
		return "";
	}
}

QString Spreadsheet::text(int row, int column) const {
	Cell *c = cell(row, column);
	if (c) {
		return c->text();
	}
	else {
		return "";
	}
}