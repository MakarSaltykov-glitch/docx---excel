//
//  mainApp.swift
//  main
//
//  Created by Makar on 02.06.2026.
//

import SwiftUI

@main
struct mainApp: App {
    var body: some Scene {
        DocumentGroup(newDocument: mainDocument()) { file in
            ContentView(document: file.$document)
        }
    }
}
