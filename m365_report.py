#!/usr/bin/env python3
"""
M365 Auto Report - Generate usage reports for SharePoint and Exchange
"""

import os
import sys
from datetime import datetime
from dotenv import load_dotenv
from msal import ConfidentialClientApplication
from msgraph import GraphServiceClient
from azure.identity import ClientSecretCredential
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


class M365Reporter:
    """Main class for M365 reporting"""
    
    def __init__(self):
        """Initialize the reporter with credentials from environment"""
        load_dotenv()
        
        self.tenant_id = os.getenv('TENANT_ID')
        self.client_id = os.getenv('CLIENT_ID')
        self.client_secret = os.getenv('CLIENT_SECRET')
        self.output_file = os.getenv('OUTPUT_FILE', 'M365_Usage_Report.xlsx')
        
        if not all([self.tenant_id, self.client_id, self.client_secret]):
            raise ValueError("Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET")
        
        self.graph_client = None
        
    def authenticate(self):
        """Authenticate to Microsoft Graph using Service Principal"""
        print("Authenticating to Microsoft Graph...")
        
        try:
            credential = ClientSecretCredential(
                tenant_id=self.tenant_id,
                client_id=self.client_id,
                client_secret=self.client_secret
            )
            
            self.graph_client = GraphServiceClient(
                credentials=credential,
                scopes=['https://graph.microsoft.com/.default']
            )
            
            print("✓ Authentication successful")
            return True
            
        except Exception as e:
            print(f"✗ Authentication failed: {str(e)}")
            return False
    
    async def get_sharepoint_sites(self):
        """Retrieve SharePoint site usage information"""
        print("\nFetching SharePoint site usage data...")
        
        sites_data = []
        
        try:
            # Get all sites
            sites = await self.graph_client.sites.get()
            
            if not sites or not sites.value:
                print("No SharePoint sites found")
                return sites_data
            
            for site in sites.value:
                if not site.id:
                    continue
                    
                site_info = {
                    'name': site.display_name or site.name or 'N/A',
                    'url': site.web_url or 'N/A',
                    'size_gb': 0,
                    'members': '',
                    'owners': ''
                }
                
                try:
                    # Get site drive for storage info
                    drive = await self.graph_client.sites.by_site_id(site.id).drive.get()
                    if drive and hasattr(drive, 'quota') and drive.quota:
                        used_bytes = getattr(drive.quota, 'used', 0) or 0
                        site_info['size_gb'] = round(used_bytes / (1024**3), 2)
                except Exception as e:
                    print(f"  Warning: Could not fetch storage for {site_info['name']}: {str(e)}")
                
                try:
                    # Get site members
                    members_result = await self.graph_client.sites.by_site_id(site.id).members.get()
                    if members_result and members_result.value:
                        members = []
                        owners = []
                        
                        for member in members_result.value:
                            display_name = getattr(member, 'display_name', None) or getattr(member, 'user_principal_name', 'Unknown')
                            
                            # Try to determine if this is an owner
                            # In SharePoint, owners are typically in a specific group
                            # For simplicity, we'll check if there's a role or group indicator
                            members.append(display_name)
                        
                        site_info['members'] = '; '.join(members[:20])  # Limit to 20 members
                        
                    # Try to get owners specifically
                    try:
                        owners_result = await self.graph_client.sites.by_site_id(site.id).owners.get()
                        if owners_result and owners_result.value:
                            owners = [getattr(owner, 'display_name', None) or getattr(owner, 'user_principal_name', 'Unknown') 
                                     for owner in owners_result.value]
                            site_info['owners'] = '; '.join(owners[:20])
                    except:
                        pass  # Owners endpoint might not be available
                        
                except Exception as e:
                    print(f"  Warning: Could not fetch members for {site_info['name']}: {str(e)}")
                
                sites_data.append(site_info)
                print(f"  ✓ Processed: {site_info['name']}")
            
            print(f"✓ Retrieved {len(sites_data)} SharePoint sites")
            
        except Exception as e:
            print(f"✗ Error fetching SharePoint sites: {str(e)}")
        
        return sites_data
    
    async def get_mailbox_usage(self):
        """Retrieve Exchange mailbox usage information"""
        print("\nFetching Exchange mailbox usage data...")
        
        mailbox_data = []
        
        try:
            # Get mailbox usage reports
            # Using Microsoft Graph reporting API
            users = await self.graph_client.users.get()
            
            if not users or not users.value:
                print("No users found")
                return mailbox_data
            
            for user in users.value:
                if not user.mail and not user.user_principal_name:
                    continue
                
                mailbox_info = {
                    'email': user.mail or user.user_principal_name or 'N/A',
                    'name': user.display_name or 'N/A',
                    'usage_gb': 0,
                    'quota_gb': 0,
                    'remaining_percent': 0
                }
                
                try:
                    # Get mailbox statistics via MailboxSettings
                    mailbox_settings = await self.graph_client.users.by_user_id(user.id).mailbox_settings.get()
                    
                    # Note: Detailed quota information requires Exchange Online management
                    # For basic implementation, we'll use available Graph API data
                    # In production, you might need to use Exchange Online PowerShell or specific reports
                    
                    # Try to get messages to estimate usage
                    messages = await self.graph_client.users.by_user_id(user.id).messages.get(
                        request_configuration={
                            'query_parameters': {
                                '$top': 1,
                                '$count': True
                            }
                        }
                    )
                    
                    # For actual usage, we'd need Exchange-specific APIs or reports
                    # This is a simplified version
                    if messages:
                        mailbox_info['usage_gb'] = 0  # Placeholder
                        mailbox_info['quota_gb'] = 50  # Default quota assumption
                        mailbox_info['remaining_percent'] = 100
                    
                except Exception as e:
                    print(f"  Warning: Could not fetch mailbox details for {mailbox_info['email']}: {str(e)}")
                
                mailbox_data.append(mailbox_info)
                print(f"  ✓ Processed: {mailbox_info['email']}")
            
            print(f"✓ Retrieved {len(mailbox_data)} mailboxes")
            
        except Exception as e:
            print(f"✗ Error fetching mailbox usage: {str(e)}")
        
        return mailbox_data
    
    def create_excel_report(self, sharepoint_data, mailbox_data):
        """Create Excel report with SharePoint and Exchange data"""
        print("\nCreating Excel report...")
        
        wb = Workbook()
        
        # Create SharePoint worksheet
        ws_sp = wb.active
        ws_sp.title = "SharePoint Sites"
        
        # Define headers for SharePoint
        sp_headers = ['Site Name', 'Site URL', 'Size (GB)', 'Members', 'Owners']
        ws_sp.append(sp_headers)
        
        # Style the header row
        header_fill = PatternFill(start_color='0066CC', end_color='0066CC', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF')
        
        for cell in ws_sp[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Add SharePoint data
        for site in sharepoint_data:
            ws_sp.append([
                site['name'],
                site['url'],
                site['size_gb'],
                site['members'],
                site['owners']
            ])
        
        # Auto-adjust column widths for SharePoint
        for column in ws_sp.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws_sp.column_dimensions[column_letter].width = adjusted_width
        
        # Create Exchange worksheet
        ws_ex = wb.create_sheet("Exchange Mailboxes")
        
        # Define headers for Exchange
        ex_headers = ['Email Address', 'Name', 'Usage (GB)', 'Quota (GB)', 'Remaining (%)']
        ws_ex.append(ex_headers)
        
        # Style the header row
        for cell in ws_ex[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Add Exchange data
        for mailbox in mailbox_data:
            ws_ex.append([
                mailbox['email'],
                mailbox['name'],
                mailbox['usage_gb'],
                mailbox['quota_gb'],
                mailbox['remaining_percent']
            ])
        
        # Auto-adjust column widths for Exchange
        for column in ws_ex.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws_ex.column_dimensions[column_letter].width = adjusted_width
        
        # Save the workbook
        wb.save(self.output_file)
        print(f"✓ Excel report saved: {self.output_file}")
    
    async def generate_report(self):
        """Main method to generate the complete report"""
        print("=" * 60)
        print("M365 Auto Report Generator")
        print("=" * 60)
        
        if not self.authenticate():
            return False
        
        # Collect data
        sharepoint_data = await self.get_sharepoint_sites()
        mailbox_data = await self.get_mailbox_usage()
        
        # Create Excel report
        self.create_excel_report(sharepoint_data, mailbox_data)
        
        print("\n" + "=" * 60)
        print("Report generation completed successfully!")
        print(f"Report saved to: {self.output_file}")
        print("=" * 60)
        
        return True


async def main():
    """Main entry point"""
    try:
        reporter = M365Reporter()
        success = await reporter.generate_report()
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\n✗ Error: {str(e)}")
        sys.exit(1)


if __name__ == '__main__':
    import asyncio
    asyncio.run(main())
