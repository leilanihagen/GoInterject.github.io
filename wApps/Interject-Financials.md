---
title: INTERJECT Financials
layout: custom
keywords: []
description: INTERJECT™ Financials specifics (Topics that are unique/specific to the Financials Application) 
---
## Overview

## My Apps
**Interject Financials:** Includes 44 data portals related to *Financials for Spreadsheets.* Data connection is redirected through "" 
**Epicor Tools** Includes 3 data portals to help migrate FRx reports. 

## REFERENCE
This deployment documents to steps and procedures to deploy Interject’s Financials for Spreadsheets as a replacement of FRx Financial Reports for Epicor Enterprise 7.x.
FRx, a Microsoft product, has come to end-of-life and official support is ending.
Even though this initial deployment focuses on the replacement of FRx and financial reporting, Interject will also be utilized across other departments, functions and data sources to bring SQL based reporting into the familiar Excel user interface.
This deployment guide assumes the migration of the FRx Sysdata has already taken place in a development environment. Databases and tools used during the FRx Sysdata migration is not needed on BETA and PROD servers and therefore excluded from this guide. 

## Media/Source Files

| Setting                    | Value                                                                                                                                                                    |
| -------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Interject Release          | Version 2.2.8.68 \| Release: Production\4.0 32x86 \| Method: Cloud\Platform                                                                                              |
| Interject Client           | Download and Install from: <br> https://install.gointerject.com/installs/Interject_Version_Installer.exe <br> https://portal.gointerject.com/kb/HowToUse/Installing.html |
| Interject Database Scripts | \\\JAXM-FILES.intuition.com\laminin\interject\software\ 