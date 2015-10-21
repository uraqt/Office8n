# -*- coding: utf-8 -*-
#
#  AppDelegate.py
#  Office8n
#
#  Created by francois on 21/10/15.
#  Copyright (c) 2015 ftiff. All rights reserved.
#

from Foundation import *
from AppKit import *
from Cocoa import *
from os.path import expanduser

import objc
import CoreFoundation

class AppDelegate(NSObject):
    
    # Setup Excel variables
    excelDropDown = objc.IBOutlet()
    excelLabel = objc.IBOutlet()
    excelID = expanduser("~") + "/Library/Containers/com.microsoft.Excel/Data/Library/Preferences/com.microsoft.Excel"
    
    # Setup Word variables
    wordDropDown = objc.IBOutlet()
    wordLabel = objc.IBOutlet()
    wordID = expanduser("~") + "/Library/Containers/com.microsoft.Word/Data/Library/Preferences/com.microsoft.Word"
    
    # Setup all languages
    availableLanguages = {
        
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
    languagesToCode = {v: k for k, v in availableLanguages.items()}
    
    
    def setupDropDown(self, _dropDown):
        _dropDown.removeAllItems()
        for languageCode, language in self.availableLanguages.items():
            _dropDown.addItemWithTitle_(language)

    def setupLabel(self, _label, _id):
        result = self.getPreference(_id, "AppleLanguages")
        currentLanguage = self.availableLanguages.get(result, "Unknown")
        _label.setStringValue_("Current Language: " + currentLanguage)

    def applicationDidFinishLaunching_(self, sender):
        NSLog("Application did finish launching.")
        
        self.setupDropDown(self.excelDropDown)
        self.setupLabel(self.excelLabel, self.excelID)

        self.setupDropDown(self.wordDropDown)
        self.setupLabel(self.wordLabel, self.wordID)

    @objc.IBAction
    def excelDropDownSelect_(self, sender):
        newLanguage = sender.titleOfSelectedItem()
        newLanguageCode = self.languagesToCode.get(newLanguage, "en")
        
        self.excelLabel.setStringValue_("Setting language to: " + newLanguage)
        self.savePreference('AppleLanguages', newLanguageCode, self.excelID, self.excelLabel)

    @objc.IBAction
    def wordDropDownSelect_(self, sender):
        newLanguage = sender.titleOfSelectedItem()
        newLanguageCode = self.languagesToCode.get(newLanguage, "en")
        
        self.wordLabel.setStringValue_("Setting language to: " + newLanguage)
        self.savePreference('AppleLanguages', newLanguageCode, self.wordID, self.wordLabel)
    
    
    def getPreference(self, appName, propertyName):
        result = CoreFoundation.CFPreferencesCopyAppValue('AppleLanguages', appName)

        # if it's an array, return only first value
        if isinstance(result, NSArray):
            result = result[0]

        return result
    
    
    def extractLanguage(self, input):
        p = re.compile(ur'["\'](.*?)["\']')
        result = ""
        try:
            result = re.search(p, input)
            result = result2.group(1)
        except ValueError:
            result = "en"
        finally:
            return result

#self.excelLabel.setStringValue_("Set to: " + propertyValue[0])
        
    
    def savePreference(self, propertyName, _value, appName, _label):
        #TODO: Change from subprocess to native call
        propertyValue = []
        propertyValue.append(_value)
        CoreFoundation.CFPreferencesSetAppValue(propertyName, propertyValue, appName)
        CoreFoundation.CFPreferencesAppSynchronize(appName)
        print "{appName}: {propertyName} = {propertyValue}".format(
                                                                   appName=appName,
                                                                   propertyName=propertyName,
                                                                   propertyValue=propertyValue
                                                                   )
        _label.setStringValue_("Set to: " + self.availableLanguages.get(propertyValue[0], "None"))
