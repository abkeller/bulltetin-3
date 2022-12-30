# -*- coding: utf-8 -*-
"""
Created on Fri Nov 13 12:56:12 2020

@author: akeller
"""

/*
by Addy Osmani
based on the original webkitdirectory POC
by Ryan Seddon
*/
var files, 
    file, 
    extension, 
    input = document.getElementById("fileURL"),
    output = document.getElementById("fileOutput"),
    holder = document.getElementById("fileHolder");

input.addEventListener("change", function (e) {
    files = e.target.files;
    output.innerHTML = "";

    for (var i = 0, len = files.length; i < len; i++) {
        file = files[i];
        extension = file.name.split(".").pop();
        output.innerHTML += "<li class='type-" + extension + "'>" + file.name + " (" +  Math.floor(file.size/1024 * 100)/100 + "KB)</li>";
    }
}, false);



// This event is fired as the mouse is moved over an element when a drag is occuring
input.addEventListener("dragover", function (e) {
    holder.classList.add("highlightOver");
});

// This event is fired when the mouse leaves an element while a drag is occuring
input.addEventListener("dragleave", function (e) {
    holder.classList.remove("highlightOver");
});

// Fires when the user releases the mouse button while dragging an object.
input.addEventListener("dragend", function (e) {
    holder.classList.remove("highlightOver");
});

// The drop event is fired on the element where the drop was occured at the end of the drag operation
input.addEventListener("drop", function (e) {
    holder.classList.remove("highlightOver");
});


