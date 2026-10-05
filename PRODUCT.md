# Product

<!-- impeccable:product-schema 1 -->

## Platform

adaptive

## Stack

delegated: Python desktop application using PySide6 because the product is a local Windows GIS review tool.

## Users

Land-data and GIS review teams who inspect polygon records before the final Excel handoff.

## Product Purpose

The application finds inconsistent, incomplete, and duplicated parcel data in KML/KMZ files, supports correction, and prevents Excel export until blocking issues are resolved.

## Positioning

It turns inconsistent GIS descriptions into structured parcel records while preserving the source KML/KMZ and providing a Google Earth review file.

## Operating Context

Users work locally with KML/KMZ land parcels, review coordinates in Google Earth, then produce an Excel deliverable.

## Capabilities and Constraints

Arabic descriptions can be plain text or HTML and use varied separators. The app operates locally, treats each Placemark with one or more polygons as one record, and writes a normalized Arabic description after edits.

## Evidence on Hand

The expected Arabic field examples and validation rules came from the requester. No source KML/KMZ sample file was supplied during implementation.

## Product Principles

- Make review blockers explicit and actionable.
- Preserve the original geographic structure whenever possible.
- Keep automatic changes safe, visible, and reversible through a separately exported file.
- Make manual correction and Google Earth matching fast.
