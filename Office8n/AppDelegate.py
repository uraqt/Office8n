# -*- coding: utf-8 -*-
#
#  AppDelegate.py
#  Office8n
#
#  Created by Francois Levaux-Tiffreau on 21/10/15.
#  Copyright (c) 2015 ftiff. All rights reserved.
#
#    This program is free software; you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation; version 2.
#    
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#    
#    You should have received a copy of the GNU General Public License along
#    with this program; if not, write to the Free Software Foundation, Inc.,
#    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.


from Foundation import *
from AppKit import *
from Cocoa import *

import objc
import CoreFoundation

class AppDelegate(NSObject):
    
    # Global Variables
    
    homeDirectory = NSHomeDirectory()
    
    # Setup Excel variables
    excelDropDown = objc.IBOutlet()
    excelLabel = objc.IBOutlet()
    excelID = homeDirectory + "/Library/Containers/com.microsoft.Excel/Data/Library/Preferences/com.microsoft.Excel"
    
    # Setup Word variables
    wordDropDown = objc.IBOutlet()
    wordLabel = objc.IBOutlet()
    wordID = homeDirectory + "/Library/Containers/com.microsoft.Word/Data/Library/Preferences/com.microsoft.Word"
    
    # Setup all languages
    codeToLanguage = {
        
                            "pt-BR": "Brazilian Portuguese",
                            "zh-CN": "Chinese (Simplified)",
                            "zh-TW": "Chinese (Traditional)",
                            "da": "Danish",
                            "nl": "Dutch",
                            "en": "English",
                            "fi": "Finnish",
                            "fr-FR": "French",
                            "de": "German",
                            "it": "Italian",
                            "ja": "Japanese",
                            "nb-NO": "Norwegian Bokmal",
                            "pl": "Polish",
                            "ru": "Russian",
                            "es": "Spanish",
                            "sv": "Swedish"
                            }

    # Invert the dictionary. Not wonderful.
    languageToCode = {v: k for k, v in codeToLanguage.items()}
    
    
    def setupDropDown(self, _dropDown):
        _dropDown.removeAllItems()
        for languageCode, language in self.codeToLanguage.items():
            _dropDown.addItemWithTitle_(language)

    def setupLabel(self, _label, _id, _dropDown):
        result = self.getPreference(_id, "AppleLanguages", _dropDown)
        currentLanguage = self.codeToLanguage.get(result, "Unknown")
        _label.setStringValue_("Current Language: " + currentLanguage)

    def applicationDidFinishLaunching_(self, sender):
        
        self.setupDropDown(self.excelDropDown)
        self.setupLabel(self.excelLabel, self.excelID, self.excelDropDown)

        self.setupDropDown(self.wordDropDown)
        self.setupLabel(self.wordLabel, self.wordID, self.wordDropDown)

    @objc.IBAction
    def excelDropDownSelect_(self, sender):
        newLanguage = sender.titleOfSelectedItem()
        newLanguageCode = self.languageToCode.get(newLanguage, "en")
        
        self.excelLabel.setStringValue_("Setting language to: " + newLanguage)
        self.savePreference('AppleLanguages', newLanguageCode, self.excelID, self.excelLabel)

    @objc.IBAction
    def wordDropDownSelect_(self, sender):
        newLanguage = sender.titleOfSelectedItem()
        newLanguageCode = self.languageToCode.get(newLanguage, "en")
        
        self.wordLabel.setStringValue_("Setting language to: " + newLanguage)
        self.savePreference('AppleLanguages', newLanguageCode, self.wordID, self.wordLabel)
    
    @objc.IBAction
    def quitButton_(self, sender):
        NSApp.terminate_(0)
    
    def getPreference(self, appName, propertyName, _dropDown):
        result = CoreFoundation.CFPreferencesCopyAppValue('AppleLanguages', appName)
        # if it's an array, return only first value
        if isinstance(result, NSArray):
            result = result[0]
        language = self.codeToLanguage.get(result, "English")
        _dropDown.selectItemWithTitle_(language)
        return result
        
    
    def savePreference(self, propertyName, _value, appName, _label):
        propertyValue = []
        propertyValue.append(_value)
        CoreFoundation.CFPreferencesSetAppValue(propertyName, propertyValue, appName)
        CoreFoundation.CFPreferencesAppSynchronize(appName)
        NSLog("{appName}: {propertyName} = {propertyValue}".format(
                                                                   appName=appName,
                                                                   propertyName=propertyName,
                                                                   propertyValue=propertyValue
                                                                   ))
        _label.setStringValue_("Set to: " + self.codeToLanguage.get(propertyValue[0], "None"))
