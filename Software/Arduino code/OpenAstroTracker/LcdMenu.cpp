#include "Utility.hpp"
#include "EPROMStore.hpp"
#include "LcdMenu.hpp"

#if HEADLESS_CLIENT == 0

// Class that drives the LCD screen with a menu
// You add a string and an id item and this class handles the display and navigation
// Create a new menu, using the given number of LCD display columns and rows
LcdMenu::LcdMenu(byte cols, byte rows, int maxItems) : _lcd(8, 9, 4, 5, 6, 7) {
  //_lcd = new LiquidCrystal(8, 9, 4, 5, 6, 7);
  _lcd.begin(cols, rows);
  _numMenuItems = 0;
  _activeMenuIndex = 0;
  _longestDisplay = 0;
  _columns = cols;
  _activeRow = -1;
  _activeCol = -1;
  _lastDisplay[0] = "";
  _lastDisplay[1] = "";
  _menuItems = new MenuItem * [maxItems];  

  _brightness = EPROMStore::Storage()->read(11);
  LOGV2(DEBUG_INFO, "LCD: Brightness from EEPROM is %d", _brightness);
  // pinMode(10, OUTPUT);
  // analogWrite(10, _brightness);

  // Create special characters for degrees and arrows
  _lcd.createChar(_degrees, DegreesBitmap);
  _lcd.createChar(_minutes, MinutesBitmap);
  _lcd.createChar(_leftArrow, LeftArrowBitmap);
  _lcd.createChar(_rightArrow, RightArrowBitmap);
  _lcd.createChar(_upArrow, UpArrowBitmap);
  _lcd.createChar(_downArrow, DownArrowBitmap);
}

// Find a menu item by its ID
MenuItem* LcdMenu::findById(byte id)
{
  for (byte i = 0; i < _numMenuItems; i++) {
    if (_menuItems[i]->id() == id) {
      return _menuItems[i];
    }
  }
  return NULL;
}

// Add a new menu item to the list (order matters)
void LcdMenu::addItem(const char* disp, byte id) {
  _menuItems[_numMenuItems++] = new MenuItem(disp, id);
  _longestDisplay = max(_longestDisplay, strlen(disp));
}

// Get the currently active item ID
byte LcdMenu::getActive() {
  return _menuItems[_activeMenuIndex]->id();
}

// Set the active menu item
void LcdMenu::setActive(byte id) {
  for (byte i = 0; i < _numMenuItems; i++) {
    if (_menuItems[i]->id() == id) {
      _activeMenuIndex = i;
      break;
    }
  }
}

// Pass thru utility function
void LcdMenu::setCursor(byte col, byte row) {
  _activeRow = row;
  _activeCol = col;
}

// Pass thru utility function
void LcdMenu::clear() {
  _lcd.clear();
}

// Set the brightness of the backlight
void LcdMenu::setBacklightBrightness(int level, bool persist) {
  _brightness = level;

  LOGV2(DEBUG_INFO, "LCD: Writing %d as brightness", _brightness  );

  // analogWrite(10, _brightness);

  LOGV2(DEBUG_INFO, "LCD: Wrote %d as brightness", _brightness  );
  if (persist) {
    LOGV2(DEBUG_INFO, "LCD: Saving %d as brightness", (_brightness & 0x00FF));
    EPROMStore::Storage()->update(11, (byte)(_brightness & 0x00FF));
  }
}

// Get the current brightness
int LcdMenu::getBacklightBrightness() {
  return _brightness;
}

// Go to the next menu item from currently active one
void LcdMenu::setNextActive() {

  _activeMenuIndex = adjustWrap(_activeMenuIndex, 1, 0, _numMenuItems - 1);

  // Update the display
  updateDisplay();

  // Clear submenu line, in case new menu doesn't print anything.
  _lcd.setCursor(0, 1);
  for (byte i = 0; i < _columns; i++) {
    _lcd.print(" ");
  }
}

// Update the display of the LCD with the current menu settings
// This iterates over the menu items, building a menu string by concatenating their display string.
// It also places the selector arrows around the active one.
// It then sends the string to the LCD, keeping the selector arrows centered in the same place.
void LcdMenu::updateDisplay() {

  char bufMenu[17];
  char* pBufMenu = &bufMenu[0];
  String menuString = "";
  byte offsetToActive = 0;
  byte offset = 0;

  char scratchBuffer[12];
  // Build the entire menu string
  for (byte i = 0; i < _numMenuItems; i++) {
    MenuItem* item = _menuItems[i];
    bool isActive = i == _activeMenuIndex;
    sprintf(scratchBuffer, "%c%s%c", isActive ? '>' : ' ', item->display(), isActive ? '<' : ' ');

    // For the active item remember where it starts in the string and insert selector arrows
    offsetToActive = isActive ? offset : offsetToActive;
    menuString += String(scratchBuffer);
    offset += strlen(scratchBuffer);
  }

  _lcd.setCursor(0, 0);
  _activeRow = 0;
  _activeCol = 0;

  // Determine where to place the active menu item. (empty space around longest item divided by two).
  int margin = (_columns - (_longestDisplay)) / 2;
  int offsetIntoString = offsetToActive - margin;

  // Pad the front if we don't have enough to offset the string to the arrow locations (happens on first item(s))
  while (offsetIntoString < 0) {
    *(pBufMenu++) = ' ';
    offsetIntoString++;
  }

  // Display the actual menu string
  while ((pBufMenu < bufMenu + _columns) && (offsetIntoString < (int)menuString.length())) {
    *(pBufMenu++) = menuString[offsetIntoString++];
  }

  // Pad the end with spaces so the display is cleared when getting to the last item(s).
  while (pBufMenu < bufMenu + _columns) {
    *(pBufMenu++) = ' ';
  }
  *(pBufMenu++) = 0;

  printMenu(String(bufMenu));

  setCursor(0, 1);
}

void LcdMenu::printChar(char ch) {
  if (ch == '>') {
    _lcd.write(_rightArrow);
  }
  else if (ch == '<') {
    _lcd.write(_leftArrow);
  }
  else if (ch == '^') {
    _lcd.write(_upArrow);
  }
  else if (ch == '~') {
    _lcd.write(_downArrow);
  }
  else if (ch == '@') {
    _lcd.write(_degrees);
  }
  else if (ch == '\'') {
    _lcd.write(_minutes);
  }
  else {
    _lcd.print(ch);
  }
}

// Print a string to the LCD at the current cursor position, substituting the special arrows and padding with spaces to the end
void LcdMenu::printMenu(String line) {
  if ((_lastDisplay[_activeRow] != line) || (_activeCol != 0)) {

    _lastDisplay[_activeRow] = line;

    _lcd.setCursor(_activeCol, _activeRow);
    int spaces = _columns - line.length();
    for (unsigned int i = 0; i < line.length(); i++) {
      printChar(line[i]);
    }

    // Clear the rest of the display
    while (spaces > 0) {
      _lcd.print(" ");
      spaces--;
    }
  }
}

// The right arrow bitmap
byte LcdMenu::RightArrowBitmap[8] = {
  B00000,
  B01000,
  B01100,
  B01110,
  B01100,
  B01000,
  B00000,
  B00000
};

// The left arrow bitmap
byte LcdMenu::LeftArrowBitmap[8] = {
  B00000,
  B00010,
  B00110,
  B01110,
  B00110,
  B00010,
  B00000,
  B00000
};

byte LcdMenu::UpArrowBitmap[8] = {
  B00100,
  B01110,
  B11111,
  B00100,
  B00100,
  B00100,
  B00100,
  B00100
};

byte LcdMenu::DownArrowBitmap[8] = {
  B000100,
  B000100,
  B000100,
  B000100,
  B000100,
  B011111,
  B001110,
  B000100
};

byte LcdMenu::DegreesBitmap[8] = {
  B01100,
  B10010,
  B10010,
  B01100,
  B00000,
  B00000,
  B00000,
  B00000
};

byte LcdMenu::MinutesBitmap[8] = {
  B01000,
  B01000,
  B01000,
  B00000,
  B00000,
  B00000,
  B00000,
  B00000
};
#else

LcdMenu::LcdMenu(byte cols, byte rows, int maxItems) {
}

MenuItem* LcdMenu::findById(byte id) {
  return NULL;
}

void LcdMenu::addItem(const char* disp, byte id) {}

byte LcdMenu::getActive() {
  return 0;
}

void LcdMenu::setActive(byte id) {}

void LcdMenu::setCursor(byte col, byte row) {}

void LcdMenu::clear() {}

void LcdMenu::setNextActive() {}

void LcdMenu::updateDisplay() {}

void LcdMenu::printMenu(String line) {}

void LcdMenu::printChar(char ch) {}

#endif

/*
class SubMenu {
	
  display() { }
  onUp() { } 
  onDown() { } 

	}
}

class Menu {
	List<SubMenu> subMenu
	int activeSubIndex=0
	UseContinuousKeys(UP|DOWN|LEFT|RIGHT)

	virtual OnUp() { previousSubmenu }
	virtual OnDown() { nextSubmenu }
	virtual OnRight() { nextmenu }	
	virtual OnLeft() { nextItemInSubmenu }
	virtual OnSelect() { confirm }
	
	displayMenu() { }
	
	run(){
		if (any continuous keys)
		{
			if 
		}
	}
}

class MenuSystem {
	List<Menu> menus;
	
	run() {
		if (!activeMenu->run())
		{
			activeMenu++
		}
	}
}

Menu RA ;
RA.fnDisplay = [] { _lcdMenu.goto(0,1); _lcdMenu.printMenu(_mount.targetRA()); }
RA.fnUp() = [] { }
*/
