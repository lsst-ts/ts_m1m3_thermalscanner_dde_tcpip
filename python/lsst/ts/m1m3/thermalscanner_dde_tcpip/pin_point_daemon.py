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
import socket
import subprocess
import sys

import win32ui  # noqa: F401

from . import __version__

# win32ui must be imported before dde
import dde  # isort: skip


class PinPointDaemon:
    """Daemon to readout PinPoint temperature values.

    Parameters
    ----------
    index : `int`
    port : `int`
        IP port for this server. If 0 then use a random port.
    simulation_mode : `int`, optional
        Simulation mode. The default is 0: do not simulate
    """

    def __init__(
        self,
        index: int,
        host: None | str,
        port: int,
        log: logging.Logger,
    ) -> None:
        self.log = log
        self._server: dde.PyDDEServer | None = None
        self._pin_point: dde.PyDDEConv | None = None

        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

        if host is None:
            host = ""
        self.port = port

        self.socket.bind((host, self.port))
        self.socket.listen(1)

    async def connect(self) -> None:
        self._server = dde.CreateServer()
        self._server.Create("Daemon")

        self._pin_point = dde.CreateConversation(self._server)
        self._pin_point.ConnectTo("PPMonitor", "System")

        # self.log.info("Requesting PPMonitor topics")
        # topics = self._pin_point.Request("Topics").split("\t")

        # self.log.debug("PPMonitor System's topics: %s", ",".join(topics))

        # system = topics[0]
        system = sys.argv[1]

        self.log.debug("Connecting to %s.", system)
        try:
            self._pin_point.ConnectTo("PPMonitor", system)
        except Exception as ex:
            self.log.error("Cannot connect to %s. The error was: %s", system, str(ex))
            sys.exit(1)

        self.log.info("Connected to %s", system)

    async def run(self) -> None:
        try:
            await self.connect()
        except Exception:
            self.log.info("Starting PinPoint monitor")
            subprocess.Popen(["C:\\Program Files (x86)\\GEC\\PinPoint\\PPMonitor.exe"])
            await asyncio.sleep(5)
            await self.connect()

        assert self._pin_point is not None

        self.scan_time = float(self._pin_point.Request("Average Scan Interval"))

        while True:
            await self.run_loop()

    async def run_loop(self) -> None:
        assert self._pin_point is not None

        self.log.info("Accepting connection on port %d.", self.port)
        connection, client_address = self.socket.accept()
        try:
            self.log.info("Client connected, client address is %s", client_address)
            while True:
                temperatures = self._pin_point.Request("Temperatures").split("\t")[:-1]
                self.log.debug("Temperatures: %s", ",".join(temperatures))

                connection.sendall(bytes(",".join(temperatures) + "\r\n", "ascii"))
                await asyncio.sleep(self.scan_time)
        except ConnectionAbortedError as ex:
            self.log.info(
                "Client connection from %s closed: %s.", client_address, str(ex)
            )
        finally:
            connection.close()


def run_pin_point_daemon() -> None:
    """Run the Pin Point Daemon."""
    logging.basicConfig(level=logging.INFO)

    log = logging.getLogger(__name__)
    log.info("Starting PinPoint Daemon %s", __version__)

    daemon = PinPointDaemon(2, None, 2222, log)
    asyncio.run(daemon.run())
