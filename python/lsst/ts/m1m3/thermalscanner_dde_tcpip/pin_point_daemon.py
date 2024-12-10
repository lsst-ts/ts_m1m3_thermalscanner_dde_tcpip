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

import argparse
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
    port : `int`
        IP port for this server. If 0 then use a random port.
    save_file : `argparse.FileType or None`
        If specified, save data to the given file.
    simulation_mode : `int`, optional
        Simulation mode. The default is 0: do not simulate
    """

    def __init__(
        self,
        ppmonitor_exe: str,
        host: None | str,
        port: int,
        save_file: None | argparse.FileType,
        ppmonitor_topic: None | str,
        log: logging.Logger,
    ) -> None:
        self.ppmonitor_exe = ppmonitor_exe
        self.ppmonitor_topic = ppmonitor_topic
        self.save_file = save_file
        self.log = log
        self._server: dde.PyDDEServer | None = None
        self._pin_point: dde.PyDDEConv | None = None

        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

        if host is None:
            host = ""
        self.port = port

        self.socket.bind((host, self.port))
        self.socket.listen(1)

        self._server = dde.CreateServer()
        self._server.Create("ThermalScannerDaemon")

        self._connection = None | socket.socket
        self._client_address = None | str

    async def connect(self) -> None:
        self._pin_point = dde.CreateConversation(self._server)

        if self.ppmonitor_topic is None:
            self._pin_point.ConnectTo("PPMonitor", "System")

            self.log.info(
                "Requesting PPMonitor topics. This probably doesn't work on Windows 11, "
                "shall work on Windows 10 and earlier."
            )
            topics = self._pin_point.Request("Topics").split("\t")

            self.log.debug("PPMonitor System's topics: %s", ",".join(topics))

            system = topics[0]
        else:
            system = self.ppmonitor_topic

        self.log.debug("Connecting to %s.", system)
        try:
            self._pin_point.ConnectTo("PPMonitor", system)
        except Exception as ex:
            raise RuntimeError(f"Cannot connect to {system}: {str(ex)}")

        self.log.info("Connected to %s", system)

    async def run(self) -> None:
        try:
            await self.connect()
        except Exception as ex:
            if self.ppmonitor_exe == "":
                raise RuntimeError(
                    "PPMonitor is not running and path to its binary was not provided, exiting."
                    "The error was: " + str(ex)
                )
            self.log.info("Starting PinPoint monitor: %s", self.ppmonitor_exe)
            subprocess.Popen([self.ppmonitor_exe])
            await asyncio.sleep(5)
            await self.connect()

        assert self._pin_point is not None

        self.scan_time = float(self._pin_point.Request("Average Scan Interval"))

        await asyncio.gather(self.telemetry_task(), self.listen_task())

    async def telemetry_task(self) -> None:
        assert self._pin_point is not None

        while True:
            temperatures = self._pin_point.Request("Temperatures").split("\t")[:-1]
            self.log.debug("Temperatures: %s", ",".join(temperatures))
            if self.save_file is not None:
                self.save_file.write(",".join(temperatures) + "\n")
                self.save_file.flush()

            if self._connection is not None:
                try:
                    self._connection.sendall(
                        bytes(",".join(temperatures) + "\r\n", "ascii")
                    )
                except ConnectionAbortedError as ex:
                    self.log.info(
                        "Client connection from %s closed: %s.",
                        self._client_address,
                        str(ex),
                    )
                    self._connection.close()
                    self._connection = None

            await asyncio.sleep(self.scan_time)

    async def listen_task(self) -> None:
        assert self._pin_point is not None

        while True:
            self.log.info("Accepting connection on port %d.", self.port)
            self._connection, self._client_address = self.socket.accept()
            self.log.info(
                "Client connected, client address is %s", self._client_address
            )
            while self._connection is not None:
                await asyncio.sleep(1)


def run_pin_point_daemon() -> None:
    """Run the Pin Point Daemon."""
    logging.basicConfig(level=logging.INFO)

    parser = argparse.ArgumentParser(
        description="Server  temperatures retrieved through DDE from the PinPoint Monitor."
    )

    default_exe = "C:\\Program Files (x86)\\GEC\\PinPoint\\PPMonitor.exe"

    parser.add_argument(
        "--ppmonitor-exe",
        default=default_exe,
        help="Location of the PinPoint Monitor exe. Default is " + default_exe,
    )

    parser.add_argument(
        "--port", default=4447, type=int, help="Port on which data will be served."
    )

    parser.add_argument(
        "--discover",
        default=False,
        type=bool,
        help=(
            "Auto discover topic in the PPPMonitor. Defaults to false, then ppmonitor-topic "
            "has to be provide."
        ),
    )

    parser.add_argument(
        "--ppmonitor-topic",
        default=None,
        type=str,
        help="PinPoint Monitor DDE topics. Equals to project, not input monitor, filename - e.g. GE01.ppc",
    )

    parser.add_argument(
        "--save",
        default=None,
        type=argparse.FileType("w"),
        help="Save telemetry to given file",
    )

    args = parser.parse_args()

    if args.discover is True:
        ppmonitor_topic = None
    else:
        if args.ppmonitor_topic is None:
            print(
                "Either --discover or --ppmonitor-topic command line argument is required"
            )
            sys.exit(1)

        ppmonitor_topic = args.ppmonitor_topic

    log = logging.getLogger(__name__)
    log.info("Starting PinPoint Daemon %s", __version__)

    daemon = PinPointDaemon(
        args.ppmonitor_exe, None, args.port, args.save, ppmonitor_topic, log
    )
    asyncio.run(daemon.run())
