#include <QtGui>
#include <qaction.h>
#include <qmenubar.h>
#include <qtoolbar.h>
#include <qlabel.h>
#include <qstatusbar.h>
#include <qmessagebox.h>
#include <qfiledialog.h>
#include <qfileinfo.h>
#include <qtablewidget.h>

#include "finddialog.h"
#include "gotocelldialog.h"
#include "mainwindow.h"
#include "sortdialog.h"
#include "spreadsheet.h"

QStringList MainWindow::recentFiles;//1.1 add-in.

MainWindow::MainWindow() {
	spreadsheet = new Spreadsheet;
	setCentralWidget(spreadsheet);
	setAttribute(Qt::WA_DeleteOnClose);//1.1 add-in. Delete the newed object when close.

	createAction();
	createMenus();
	createContextMenu();
	createToolBars();
	createStatusBar();

	readSettings();

	findDialog = 0;

	setWindowIcon(QIcon(":/images/icon.png"));
	setCurrentFile("");
	

}

void MainWindow::closeEvent(QCloseEvent *event) {
	if (okToContinue()) {
		writeSettings();
		event->accept();
	}
	else {
		event->ignore();
	}
}


void MainWindow::newFile() {
	//version 1.0
	/*if (okToContinue()) {
	spreadsheet->clear();
	setCurrentFile("");
	}*/

	//version 1.1
	MainWindow *mainWin = new MainWindow;
	mainWin->show();
}

void MainWindow::open() {
	if (okToContinue()) {
		QString fileName = QFileDialog::getOpenFileName(this,
			tr("Open MySpreadsheet"), ".",
			tr("Spreadsheet Files(*.sp)"));
		if (!fileName.isEmpty())
			loadFile(fileName);
	}
}

bool MainWindow::save() {
	if (curFile.isEmpty()) {
		return saveAs();
	}
	else {
		return saveFile(curFile);
	}
}

bool MainWindow::saveAs() {
	//QString fileName = QFileDialog::getSaveFileName(this,
		//tr("Save Spreadsheet"), tr("Spreadsheet (*.sp)"));
	QString fileName =  QFileDialog::getSaveFileName(this, 
		tr("Save MySpreadsheet"), "./Untitled.sp", tr("Spreadshee Files(*.sp)"));
	//If the surffix of the fileName is empty
	if (QFileInfo(fileName).suffix().isEmpty()) 
		fileName.append(".sp");
	if (fileName.isEmpty())
		return false;

	return saveFile(fileName);
}

void MainWindow::find() {
	if (!findDialog) {
		findDialog = new FindDialog(this);
		connect(findDialog, SIGNAL(findNext(const QString&, Qt::CaseSensitivity)),
			spreadsheet, SLOT(findNext(const QString&, Qt::CaseSensitivity)));
		connect(findDialog, SIGNAL(findPrevious(const QString&, Qt::CaseSensitivity)),
			spreadsheet, SLOT(findPrevious(const QString&, Qt::CaseSensitivity)));
	}
	//If findDialog already existed, function show() won't do anything.
	findDialog->show();
	findDialog->raise();
	findDialog->activateWindow();
}

void MainWindow::goToCell() {
	GoToCellDialog dialog(this);
	if (dialog.exec()) {
		QString str = dialog.lineEdit->text().toUpper();
		spreadsheet->setCurrentCell(str.mid(1).toInt() - 1,
			str[0].unicode() - 'A');
	}
}

void MainWindow::sort() {
	SortDialog dialog(this);
	QTableWidgetSelectionRange range = spreadsheet->selectedRange();
	dialog.setColumnRange('A' + range.leftColumn(),
		'A' + range.rightColumn());

	if (dialog.exec()) {
		SpreadsheetCompare compare;
		compare.keys[0] =
			dialog.primaryColumnCombo->currentIndex();
		compare.keys[1] =
			dialog.secondaryColumnCombo->currentIndex() - 1;
		compare.keys[2] =
			dialog.tertiaryColumnCombo->currentIndex() - 1;
		compare.ascending[0] =
			(dialog.primaryColumnCombo->currentIndex() == 0);
		compare.ascending[1] =
			(dialog.secondaryColumnCombo->currentIndex() == 0);
		compare.ascending[2] =
			(dialog.tertiaryColumnCombo->currentIndex() == 0);
		spreadsheet->sort(compare);
	}
}

void MainWindow::about() {
	QMessageBox::about(this, tr("About MySpreadsheet"),
		tr("<h2>MySpreadsheet 1.1<h2>"
		"<p>MySpreadsheet is a small application that "
		"demonstrasts QAction, QMainWindow, QMenuBar, "
		"QStatusBar, QTableWidget, QToolBar, and many other"
		"Qt classes."));
}


void MainWindow::openRecentFile() {
	if (okToContinue()) {
		QAction *action = qobject_cast<QAction*>(sender());
		if (action)
			loadFile(action->data().toString());
	}
}

void MainWindow::updateStatusBar() {
	locationlabel->setText(spreadsheet->currentLocation());
	formulaLabel->setText(spreadsheet->currentFormula());
}

void MainWindow::spreadsheetModified() {
	setWindowModified(true);
	updateStatusBar();
}

void MainWindow::createAction() {
	newAction = new QAction(tr("&New"), this);
	newAction->setIcon(QIcon(":/images/new.png"));
	newAction->setShortcut(QKeySequence::New);
	newAction->setStatusTip(tr("Create a new spreadsheet file"));
	connect(newAction, SIGNAL(triggered()), this, SLOT(newFile()));

	openAction = new QAction(tr("&Open..."), this);
	openAction->setIcon(QIcon(":/images/open.png"));
	openAction->setShortcut(QKeySequence::Open);
	openAction->setStatusTip(tr("Open an existing spreadsheet file"));
	connect(openAction, SIGNAL(triggered()), this, SLOT(open()));

	saveAction = new QAction(tr("&Save"), this);
	saveAction->setIcon(QIcon(":/images/save.png"));
	saveAction->setShortcuts(QKeySequence::Save);
	saveAction->setStatusTip(tr("Save the spreadsheet to disk"));
	connect(saveAction, SIGNAL(triggered()), this, SLOT(save()));

	saveAsAction = new QAction(tr("Save &as...."), this);
	saveAsAction->setStatusTip(tr("Save the spreadsheet under a new name"));
	connect(saveAsAction, SIGNAL(triggered()), this, SLOT(saveAs()));

	for (int i = 0; i != MaxRecentFiles; ++i) {
		recentFileActions[i] = new QAction(this);
		recentFileActions[i]->setVisible(false);
		connect(recentFileActions[i], SIGNAL(triggered()), this, SLOT(openRecentFile()));
	}

	//Strat quote.
	//Add-in in version 1.1
	closeAction = new QAction(tr("&Close"), this);
	closeAction->setShortcut(QKeySequence::Close);
	closeAction->setStatusTip(tr("Close this window"));
	connect(closeAction, SIGNAL(triggered()), this, SLOT(close()));
	//End quote.



	exitAction = new QAction(tr("&Exit"), this);
	exitAction->setShortcut(tr("Ctrl+Q"));
	exitAction->setStatusTip(tr("Exit the application"));
	//version 1.0:
	//connect(exitAction, SIGNAL(triggered()), this, SLOT(close()));
	//version 1.1:
	connect(exitAction, SIGNAL(triggered()), qApp, SLOT(closeAllWindows()));

	cutAction = new QAction(tr("&Cut"), this);
	cutAction->setIcon(QIcon(":/images/cut.png"));
	cutAction->setShortcut(QKeySequence::Cut);
	cutAction->setStatusTip(tr("Cut the current selection contents to the clipboard"));
	connect(cutAction, SIGNAL(triggered()), spreadsheet, SLOT(cut()));

	copyAction = new QAction(tr("&Copy"), this);
	copyAction->setIcon(QIcon(":/images/copy.png"));
	copyAction->setShortcut(QKeySequence::Copy);
	copyAction->setStatusTip(tr("Copy the current selection's contents to the clipboard"));
	connect(copyAction, SIGNAL(triggered()), spreadsheet, SLOT(copy()));

	pasteAction = new QAction(tr("&Paste"), this);
	pasteAction->setIcon(QIcon(":/images/paste.png"));
	pasteAction->setShortcut(QKeySequence::Paste);
	pasteAction->setStatusTip(tr("Paste the clipboard's contents into the current selection"));
	connect(pasteAction, SIGNAL(triggered()), spreadsheet, SLOT(paste()));

	deleteAction = new QAction(tr("&Delete"), this);
	deleteAction->setIcon(QIcon(":/images/delete.png"));
	deleteAction->setShortcut(QKeySequence::Delete);
	deleteAction->setStatusTip(tr("Delete the current selection's contents"));
	connect(deleteAction, SIGNAL(triggered()), spreadsheet, SLOT(del()));

	selectRowAction = new QAction(tr("&Row"), this);
	selectRowAction->setStatusTip(tr("Select all the cells in the current row"));
	connect(selectRowAction, SIGNAL(triggered()), spreadsheet, SLOT(selectCurrentRow()));

	selectColumnAction = new QAction(tr("&Column"), this);
	selectColumnAction->setStatusTip(tr("select all the cells in the current column"));
	connect(selectColumnAction, SIGNAL(triggered()), spreadsheet, SLOT(selectCurrentColumn()));

	selectAllAction = new QAction(tr("&All"), this);
	selectAllAction->setShortcut(QKeySequence::SelectAll);
	selectAllAction->setStatusTip(tr("Select alll the cells in the spreadsheet"));
	connect(selectAllAction, SIGNAL(triggered()), spreadsheet, SLOT(selectAll()));

	findAction = new QAction(tr("&Find"), this);
	findAction->setIcon(QIcon(":/images/find.png"));
	findAction->setShortcut(QKeySequence::Find);
	findAction->setStatusTip(tr("Find a matching cell"));
	connect(findAction, SIGNAL(triggered()), this, SLOT(find()));

	goToCellAction = new QAction(tr("&Go to cell..."), this);
	goToCellAction->setIcon(QIcon(":/images/gotocell.png"));
	goToCellAction->setShortcut(tr("Ctrl+G"));
	goToCellAction->setStatusTip(tr("Go to the specified cell"));
	connect(goToCellAction, SIGNAL(triggered()), this, SLOT(goToCell()));

	recalculateAction = new QAction(tr("&Recalculate"), this);
	recalculateAction->setShortcut(tr("F9"));
	recalculateAction->setStatusTip(tr("Recalculate all the spreadsheet's formulas"));
	connect(recalculateAction, SIGNAL(triggered()), spreadsheet, SLOT(recalculate()));

	sortAction = new QAction(tr("&Sort.."), this);
	sortAction->setStatusTip(tr("Sort the selected cells or all the cells"));
	connect(sortAction, SIGNAL(triggered()), this, SLOT(sort()));

	showGridAction = new QAction(tr("&Show Grid"), this);
	showGridAction->setCheckable(true);
	showGridAction->setChecked(spreadsheet->showGrid());
	showGridAction->setStatusTip(tr("Show or hide the spreadshhet's grid"));
	connect(showGridAction, SIGNAL(toggled(bool)), spreadsheet, SLOT(setShowGrid(bool)));

	autoRecalcAtion = new QAction(tr("&Auto-Recalculate"), this);
	autoRecalcAtion->setCheckable(true);
	autoRecalcAtion->setChecked(spreadsheet->autoRecalculate());
	autoRecalcAtion->setStatusTip(tr("Switch auto-recalculation on or off"));
	connect(autoRecalcAtion, SIGNAL(toggled(bool)), spreadsheet, SLOT(setAutoRecalculate(bool)));


	aboutAction = new QAction(tr("&About"), this);
	aboutAction->setStatusTip(tr("Show the application's About box"));
	connect(aboutAction, SIGNAL(triggered()), this, SLOT(about()));

	aboutQtAction = new QAction(tr("About &Qt"), this);
	aboutQtAction->setStatusTip(tr("Show the Qt library's About box"));
	connect(aboutQtAction, SIGNAL(triggered()), qApp, SLOT(aboutQt()));
}

void MainWindow::createMenus() {
	fileMenu = menuBar()->addMenu(tr("&File"));
	fileMenu->addAction(newAction);
	fileMenu->addAction(openAction);
	fileMenu->addAction(saveAction);
	fileMenu->addAction(saveAsAction);
	separatorAction = fileMenu->addSeparator();
	for (int i = 0; i != MaxRecentFiles; ++i)
		fileMenu->addAction(recentFileActions[i]);
	fileMenu->addSeparator();
	fileMenu->addAction(closeAction);//Add in in version 1.1
	fileMenu->addAction(exitAction);

	editMenu = menuBar()->addMenu(tr("&Edit"));
	editMenu->addAction(cutAction);
	editMenu->addAction(copyAction);
	editMenu->addAction(pasteAction);
	editMenu->addAction(deleteAction);

	selectSubMenu = editMenu->addMenu(tr("&Select"));
	selectSubMenu->addAction(selectRowAction);
	selectSubMenu->addAction(selectColumnAction);
	selectSubMenu->addAction(selectAllAction);

	editMenu->addSeparator();
	editMenu->addAction(findAction);
	editMenu->addAction(goToCellAction);

	toolsMenu = menuBar()->addMenu(tr("&Tools"));
	toolsMenu->addAction(recalculateAction);
	toolsMenu->addAction(sortAction);

	optionsMenu = menuBar()->addMenu(tr("&Options"));
	optionsMenu->addAction(showGridAction);
	optionsMenu->addAction(autoRecalcAtion);

	menuBar()->addSeparator();

	helpMenu = menuBar()->addMenu(tr("&Help"));
	helpMenu->addAction(aboutAction);
	helpMenu->addAction(aboutQtAction);
}

//Create the menu of the right mouse button
void MainWindow::createContextMenu() {
	spreadsheet->addAction(cutAction);
	spreadsheet->addAction(copyAction);
	spreadsheet->addAction(pasteAction);
	spreadsheet->setContextMenuPolicy(Qt::ActionsContextMenu);
}

void MainWindow::createToolBars() {
	fileToolBar = addToolBar(tr("&File"));
	fileToolBar->addAction(newAction);
	fileToolBar->addAction(openAction);
	fileToolBar->addAction(saveAction);

	editToolBar = addToolBar(tr("&Edit"));
	editToolBar->addAction(cutAction);
	editToolBar->addAction(copyAction);
	editToolBar->addAction(pasteAction);
	editToolBar->addSeparator();
	editToolBar->addAction(findAction);
	editToolBar->addAction(goToCellAction);
}

void MainWindow::createStatusBar() {
	locationlabel = new QLabel(" W999 ");
	locationlabel->setAlignment(Qt::AlignHCenter);
	locationlabel->setMinimumSize(locationlabel->sizeHint());

	formulaLabel = new QLabel;
	formulaLabel->setIndent(3);

	statusBar()->addWidget(locationlabel);
	statusBar()->addWidget(formulaLabel, 1);

	//Connect the change of the selected cell' positon and the statusbar 
	connect(spreadsheet, SIGNAL(currentCellChanged(int, int, int, int)),
		this, SLOT(updateStatusBar()));
	//Connect the change of the selected cell' text and the statusbar
	connect(spreadsheet, SIGNAL(modified()), this, SLOT(spreadsheetModified()));

	updateStatusBar();
}

void MainWindow::readSettings() {
	QSettings settings("Software Inc", "Spreadsheet");

	restoreGeometry(settings.value("geometry").toByteArray());

	recentFiles = settings.value("recentFiles").toStringList();
	//Version 1.0
	//updateRecentFileActions();
	//Start note.
	//Version 1.1
	foreach(QWidget *win, QApplication::topLevelWidgets()) {
		if (MainWindow *mainWin = qobject_cast<MainWindow *>(win))
			mainWin->updateRecentFileActions();
	}
	//End note.

	bool showGrid = settings.value("showGrid", true).toBool();
	showGridAction->setChecked(showGrid);

	bool autoRecalc = settings.value("autoRecalc", true).toBool();
	autoRecalcAtion->setChecked(autoRecalc);
}

void MainWindow::writeSettings() {
	QSettings settings("Software Inc", "Spreadsheet");

	settings.setValue("geometry", saveGeometry());
	settings.setValue("recentFiles", recentFiles);
	settings.setValue("showGrid", showGridAction->isChecked());
	settings.setValue("autoRecalc", autoRecalcAtion->isChecked());
};


bool MainWindow::okToContinue() {
	if (isWindowModified()) {
		int r = QMessageBox::warning(this, tr("MySpreadsheet"),
			tr("The document has been modified.\n"
			"Do you want to save your changes?"),
			QMessageBox::Yes | QMessageBox::No | QMessageBox::Cancel);
		if (r == QMessageBox::Yes) {
			return save();
		}
		else if (r == QMessageBox::Cancel) {
			return false;
		}
	}//if unmodified or no, return ture. 
	return true;
}

bool MainWindow::loadFile(const QString &fileName) {
	if (!spreadsheet->readFile(fileName)) {
		statusBar()->showMessage(tr("Loading canceled"), 2000);
		return false;
	}

	setCurrentFile(fileName);
	statusBar()->showMessage(tr("File loaded"), 2000);
	return true;
}

bool MainWindow::saveFile(const QString &fileName) {
	if (!spreadsheet->writeFile(fileName)) {
		statusBar()->showMessage(tr("Saving canceled"), 2000);
		return false;
	}

	setCurrentFile(fileName);
	statusBar()->showMessage(tr("File saved"), 2000);
	return true;
}

void MainWindow::setCurrentFile(const QString &fileName) {
	curFile = fileName;
	setWindowModified(false);

	QString shownName = tr("Untitle");
	if (!curFile.isEmpty()) {
		shownName = strippedName(curFile);
		recentFiles.removeAll(curFile);
		recentFiles.append(curFile);
		//Version 1.0
		//updateRecentFileActions();
		//Strat note.
		//Version 1.1
		foreach(QWidget *win, QApplication::topLevelWidgets()) {
			if (MainWindow *mainWin = qobject_cast<MainWindow *>(win))
				mainWin->updateRecentFileActions();
		}
		//End note.
	}

	setWindowTitle(tr("%1[*]-%2").arg(shownName).arg(tr("MySpreadsheet")));
}

void MainWindow::updateRecentFileActions() {
	QMutableStringListIterator i(recentFiles);

	while (i.hasNext()) {
		if (!QFile::exists(i.next()))
			i.remove();
	}

	for (int j = 0; j != MaxRecentFiles; ++j) {
		if (j < recentFiles.count()) {
			//Version 1.0
			/*QString text = tr("&%1 %2")
			.arg(j + 1)
			.arg(strippedName(recentFiles[j]));
			recentFileActions[j]->setText(text);
			recentFileActions[j]->setData(recentFiles[j]);
			recentFileActions[j]->setVisible(true);*/
			//Start note.
			//Version 1.1: transpore rencentfile.
			const int &&fileIndex = recentFiles.count() - 1 - j;
			QString text = tr("&%1 %2")
				.arg(j + 1)
				.arg(strippedName(recentFiles[fileIndex]));
			recentFileActions[j]->setText(text);
			recentFileActions[j]->setData(recentFiles[fileIndex]);
			recentFileActions[j]->setVisible(true);
			//End note.
		}
		else {
			recentFileActions[j]->setVisible(false);
		}
	}
	separatorAction->setVisible(!recentFiles.isEmpty());
}

QString MainWindow::strippedName(const QString &fullFileName) {
	return QFileInfo(fullFileName).fileName();
}
