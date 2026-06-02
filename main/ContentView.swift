//
//  ContentView.swift
//  main
//
//  Created by Makar on 02.06.2026.
//

import SwiftUI

struct ContentView: View {
    @Binding var document: mainDocument

    var body: some View {
        TextEditor(text: $document.text)
    }
}

#Preview {
    ContentView(document: .constant(mainDocument()))
}
