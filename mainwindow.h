#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <qmainwindow.h>

class QAction;
class QLabel;
class FindDialog;
class Spreadsheet;

class MainWindow : public  QMainWindow
{
	Q_OBJECT;

public:
	MainWindow();

protected:
	void closeEvent(QCloseEvent *event);

	private slots:
	void newFile();
	void open();
	bool save();
	bool saveAs();
	void find();
	void goToCell();
	void sort();
	void about();
	void openRecentFile();
	void updateStatusBar();
	void spreadsheetModified();

private:
	void createAction();
	void createMenus();
	void createContextMenu();//Create the menu of the right mouse button.
	void createToolBars();
	void createStatusBar();
	void readSettings();
	void writeSettings();
	bool okToContinue();
	bool loadFile(const QString &fileName);
	bool saveFile(const QString &fileName);
	void setCurrentFile(const QString &fileName);
	void updateRecentFileActions();
	QString strippedName(const QString &fullName);

	Spreadsheet *spreadsheet;
	FindDialog *findDialog;
	QLabel *locationlabel;
	QLabel *formulaLabel;
	static QStringList recentFiles;//1.1 add-in.
	QString curFile;

	enum { MaxRecentFiles = 5 };
	QAction *recentFileActions[MaxRecentFiles];
	QAction *separatorAction;

	QMenu *fileMenu;
	QMenu *editMenu;
	QMenu *selectSubMenu;
	QMenu *toolsMenu;
	QMenu *optionsMenu;
	QMenu *helpMenu;
	QToolBar *fileToolBar;
	QToolBar *editToolBar;
	QAction *newAction;
	QAction *openAction;
	QAction *saveAction;
	QAction *saveAsAction;
	QAction *closeAction;//Added in in version 1.1
	QAction *exitAction;
	QAction *cutAction;
	QAction *copyAction;
	QAction *pasteAction;
	QAction *deleteAction;
	QAction *selectRowAction;
	QAction *selectColumnAction;
	QAction *selectAllAction;
	QAction *findAction;
	QAction *goToCellAction;
	QAction *recalculateAction;
	QAction *sortAction;
	QAction *showGridAction;
	QAction *autoRecalcAtion;
	QAction *aboutAction;
	QAction *aboutQtAction;
};



#endif