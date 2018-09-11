# -*- coding: utf-8 -*-
"""
/***************************************************************************
 miniprojet
                                 A QGIS plugin
 miniprojet
                             -------------------
        begin                : 2018-05-21
        copyright            : (C) 2018 by hanaa khoj et khalifi mohamed
        email                : hanaa.khoj@gmail.com
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""


# noinspection PyPep8Naming
def classFactory(iface):  # pylint: disable=invalid-name
    """Load miniprojet class from file miniprojet.

    :param iface: A QGIS interface instance.
    :type iface: QgisInterface
    """
    #
    from .miniprojet import miniprojet
    return miniprojet(iface)
