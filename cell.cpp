#include "cell.h"

Cell::Cell() {
	setDirty();
}

//When user use QTableWidget's Item, Item's function clone() will be called.
//In spreadsheet.cpp, Item was translated into cell.
//So cell's function clone will be called.
QTableWidgetItem *Cell::clone() const {
	return new Cell(*this);
}

void Cell::setData(int role, const QVariant &value) {
	QTableWidgetItem::setData(role, value);
	if (role == Qt::EditRole)
		setDirty();
}

QVariant Cell::data(int role) const {
	if (role == Qt::DisplayRole) { //If data's type is string(formular). 
		if (value().isValid()) {//Function value() can charge data's type.
			return value().toString();
		}
		else {
			return "####";
		}
	}
	else if (role == Qt::TextAlignmentRole) { //If data is alignment.
		if (value().type() == QVariant::String) {
			return int(Qt::AlignLeft | Qt::AlignVCenter);
		}
		else {
			return int(Qt::AlignRight | Qt::AlignVCenter);
		}
	}
	else {  // If data's type is double.
		return QTableWidgetItem::data(role);
	}
}

void Cell::setFormula(const QString &formula) {
	setData(Qt::EditRole, formula);
}

QString Cell::formula() const {
	return data(Qt::EditRole).toString();
}

void Cell::setDirty() {
	cachIsDirty = true;
}

const QVariant Invalid;


//Return data' value, which may be a double number or a string.
QVariant Cell::value() const {
	if (cachIsDirty) {
		cachIsDirty = false;

		QString formulaStr = formula();
		if (formulaStr.startsWith('\'')) { //Data in form like'12.33900.
			cachedValue = formulaStr.mid(1);
		}
		else if (formulaStr.startsWith('=')) { //Data may be a formular.
			cachedValue = Invalid;
			QString expr = formulaStr.mid(1);
			expr.replace(" ", "");
			expr.append(QChar::Null);

			int pos = 0;
			cachedValue = evalExpression(expr, pos); //Result's type should be double.
			if (expr[pos] != QChar::Null) //Eval didn't go through the whole expression.
				cachedValue = Invalid;
		}
		else { //Data's type is double or string.
			bool ok;
			double d = formulaStr.toDouble(&ok);
			if (ok) { //Data's type is double.
				cachedValue = d;
			}
			else { // Data's type is string.
				cachedValue = formulaStr;
			}
		}
	}
	return cachedValue;
}

QVariant Cell::evalExpression(const QString &str, int &pos) const {
	QVariant result = evalTerm(str, pos); //Eval the firts term.
	while (str[pos] != QChar::Null) {
		QChar op = str[pos];
		if (op != '+' && op != '-')
			return result; 
		++pos;

		QVariant term = evalTerm(str, pos); //Eval the second term.
		if (result.type() == QVariant::Double
			&& term.type() == QVariant::Double) {
			if (op == '+') {
				result = result.toDouble() + term.toDouble();
			}
			else {
				result = result.toDouble() - term.toDouble();
			}
		}
		else {
			result = Invalid;
		}
	} //End while.
	return result;
}

QVariant Cell::evalTerm(const QString &str, int &pos) const {
	QVariant result = evalFactor(str, pos); //Eval the first factor.
	while (str[pos] != QChar::Null) {
		QChar op = str[pos];
		if (op != '*' && op != '/')
			return result;
		++pos;

		QVariant factor = evalFactor(str, pos);//Eval the second factor.
		if (result.type() == QVariant::Double
			&& factor.type() == QVariant::Double) {
			if (op == '*') {
				result = result.toDouble() * factor.toDouble();
			}
			else {
				if (factor.toDouble() == 0.0) {
					result == Invalid;
				}
				else {
					result = result.toDouble() / factor.toDouble();
				}
			}
		}
		else {
			result = Invalid;
		}
	}//End while.
	return result;
}

QVariant Cell::evalFactor(const QString &str, int &pos) const {
	QVariant result;
	bool negetive = false;

	if (str[pos] == '-') {
		negetive = true;
		++pos;
	}

	if (str[pos] == '(') {
		++pos;
		result = evalExpression(str, pos);
		if (str[pos] != ')')
			result = Invalid;
		++pos;
	}
	else {
		QRegExp regExp("[A-Za-z][1-9][0-9]{0,2}");
		QString token;

		while (str[pos].isLetterOrNumber() || str[pos] == '.') {
			token += str[pos];
			++pos;
		}

		if (regExp.exactMatch(token)) { //If the factor is a positon.
			int column = token[0].toUpper().unicode() - 'A';
			int row = token.mid(1).toInt() - 1;

			Cell *c = static_cast<Cell *>(
				tableWidget()->item(row, column));
			if (c) {
				result = c->value();
			}
			else {
				result = 0.0;
			}
		}
		else { // If the factor may be a number.
			bool ok;
			result = token.toDouble(&ok);
			if (!ok) 
				result = Invalid;
		}
	}
	if (negetive) {
		if (result.type() == QVariant::Double) {
			result = -result.toDouble();
		}
		else {
			result = Invalid;
		}
	}
	return result;
}



















