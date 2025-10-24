<#
.SYNOPSIS
    Configure Windows Firewall to allow ONLY ICMP (Ping) traffic to a Windows 10 machine.

.DESCRIPTION
    This PowerShell script (and accompanying notes) describe how to enable inbound ICMP 
    Echo Requests (ping) in Windows Firewall on Windows 10. 
    ICMP (Internet Control Message Protocol) is used for diagnostic and network testing 
    purposes such as the 'ping' command.

    By default, Windows Firewall blocks unsolicited inbound connections including ICMP.
    The rules created below explicitly allow only ICMP Echo Requests (type 8 for IPv4, 
    type 128 for IPv6) to pass through.

    You can enable this using either:
    - Windows Firewall with Advanced Security (GUI)
    - PowerShell commands
    - The netsh command-line utility

    These rules can be restricted further to specific IP addresses or profiles 
    (Domain, Private, Public) as required.

====================================================================
=                  WINDOWS FIREWALL GUI STEPS                      =
====================================================================
1. Open Start → type "Windows Defender Firewall with Advanced Security" → open it.
2. In the left pane, select "Inbound Rules".
3. In the right pane, click "New Rule...".
4. In the wizard:
   a. Select "Custom" → Next.
   b. For Program → choose "All programs" → Next.
   c. For Protocol type → select "ICMPv4".
   d. Click "Customize..." → Select "Specific ICMP types" → check "Echo Request" → OK → Next.
   e. For Scope → choose Any IP address (or specify certain IPs) → Next.
   f. For Action → select "Allow the connection" → Next.
   g. Choose network profiles (Domain/Private/Public) → usually uncheck Public → Next.
   h. Name the rule (e.g., "Allow ICMPv4 Echo Request (Ping)") → Finish.
5. Repeat for "ICMPv6" if IPv6 ping should also be allowed.
6. Verify that the new rule(s) are enabled.

To remove or disable:
   - Right-click the rule → "Delete" or "Disable".

====================================================================
=                POWERSHELL COMMAND-LINE METHOD                    =
====================================================================
# Run PowerShell as Administrator before executing these commands.

# Allow inbound ICMPv4 Echo Requests (ping type 8)
New-NetFirewallRule -DisplayName "Allow ICMPv4 Echo In" `
                    -Protocol ICMPv4 -IcmpType 8 `
                    -Direction Inbound -Action Allow `
                    -Profile Any

# Allow inbound ICMPv6 Echo Requests (ping type 128)
New-NetFirewallRule -DisplayName "Allow ICMPv6 Echo In" `
                    -Protocol ICMPv6 -IcmpType 128 `
                    -Direction Inbound -Action Allow `
                    -Profile Any

# Remove the rules when no longer needed:
Remove-NetFirewallRule -DisplayName "Allow ICMPv4 Echo In"
Remove-NetFirewallRule -DisplayName "Allow ICMPv6 Echo In"

====================================================================
=                   NETSH COMMAND-LINE METHOD                      =
====================================================================
# Allow ICMPv4 Echo Requests
netsh advfirewall firewall add rule name="Allow ICMPv4 In" `
    protocol=icmpv4:8,any dir=in action=allow enable=yes profile=any

# Allow ICMPv6 Echo Requests
netsh advfirewall firewall add rule name="Allow ICMPv6 In" `
    protocol=icmpv6:128,any dir=in action=allow enable=yes profile=any

# Delete the rules
netsh advfirewall firewall delete rule name="Allow ICMPv4 In"
netsh advfirewall firewall delete rule name="Allow ICMPv6 In"

====================================================================
=                         TESTING STEPS                            =
====================================================================
1. From another system, run:
       ping <your-machine-IP>
2. If the ping replies successfully, the rule is working.
3. If it fails:
   - Confirm the firewall rule is enabled and applies to the correct profile.
   - Ensure no router or external firewall is blocking ICMP.
   - Confirm no third-party antivirus/firewall is interfering.

====================================================================
=                         SECURITY NOTES                           =
====================================================================
- ICMP Echo is generally low-risk but can expose your system to ping sweeps.
- Limit rule scope to known IPs if possible:
      New-NetFirewallRule -DisplayName "Allow ICMPv4 Echo In from AdminPC" `
                          -Protocol ICMPv4 -IcmpType 8 -Direction Inbound `
                          -Action Allow -RemoteAddress 192.168.1.10
- Avoid enabling this rule on the "Public" profile unless absolutely necessary.
- Disable or remove the rule when testing is complete.

====================================================================
=                           REFERENCES                             =
====================================================================
Microsoft Docs:
  - https://learn.microsoft.com/en-us/powershell/module/netsecurity/new-netfirewallrule
  - https://learn.microsoft.com/en-us/windows/security/threat-protection/windows-firewall/create-inbound-outbound-rules
#>


New-NetFirewallRule -DisplayName "Allow ICMPv4 Echo In" -Protocol ICMPv4 -IcmpType 8 -Direction Inbound -Action Allow -Profile Any
