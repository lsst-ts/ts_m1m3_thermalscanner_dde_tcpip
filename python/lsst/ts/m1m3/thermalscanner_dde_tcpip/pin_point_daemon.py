# This file is part of ts_m1m3_thermalscanner_dde_tcpip.
#
# Developed for the Vera Rubin Observatory Telescope and Site Systems.
# This product includes software developed by the LSST Project
# (https://www.lsst.org).
# See the COPYRIGHT file at the top-level directory of this distribution
# for details of code ownership.
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.

__all__ = ["PinPointDaemon", "run_pin_point_daemon"]

import asyncio
import logging

import win32ui  # noqa: F401

from . import __version__

# win32ui must be imported before dde
import dde  # isort: skip


logger = logging.getLogger(__name__)


class PinPointDaemon:
    """Daemon to readout PinPoint temperature values.

    Parameters
    ----------
    index : `int`
    """

    def __init__(self) -> None:
        self._server: dde.PyDDEServer | None = None
        self._pin_point: dde.PyDDEConv | None = None

    def connect(self) -> None:
        self._server = dde.CreateServer()
        self._server.Create("Daemon")

        self._pin_point = dde.CreateConversation(self._server)
        self._pin_point.ConnectTo("PPMonitor", "System")

        logger.info("Requesting PPMonitor topics")
        topics = self._pin_point.Request("Topics").split("\t")

        logger.debug("PPMonitor System's topics: %s", ",".join(topics))

        system = topics[0]

        self._pin_point.ConnectTo("PPMonitor", system)

        logger.info("Connected to %s", system)

    async def run_loop(self) -> None:
        assert self._pin_point is not None

        scan_time = float(self._pin_point.Request("Average Scan Interval"))
        while True:
            temperatures = self._pin_point.Request("Temperatures").split("\t")[:-1]
            logger.info("Temperatures: %s", ",".join(temperatures))

            temp = [float(s) for s in temperatures]
            print(len(temp), temp)

            await asyncio.sleep(scan_time)


def run_pin_point_daemon() -> None:
    """Run the Pin Point Daemon."""
    logging.basicConfig(level=logging.INFO)
    logger.info("Starting PinPoint Daemon %s", __version__)

    daemon = PinPointDaemon()
    daemon.connect()

    asyncio.run(daemon.run_loop())
